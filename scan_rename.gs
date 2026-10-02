const SOURCE_FOLDER_ID = '1Oln0vKLfmNInPyHVMPnR22tEtp9KXozB'; // マイドライブ/scan/未処理データ
const DEST_FOLDER_ID   = '1bTo9kY4ZJQCtMeURLJZEHOTIHoyfWJJN'; // マイドライブ/scan/処理済みデータ

// PDFの回転・画像→PDF変換に使用するライブラリ（pdf-lib）
const PDF_LIB_URL = 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js';

// 画像をPDF化する際のページサイズ（長辺をA4相当 842pt に合わせる）
const PAGE_LONG_SIDE_PT = 842;

async function checkForNewFiles() {
  const startTime = Date.now(); // 実行開始時間を記録
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000; // 4.5分（270,000ミリ秒）を上限とする

  // 二重起動防止（前回の処理が実行中なら今回はスキップ）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log("前回の処理が実行中のため、今回はスキップします。");
    return;
  }

  try {
    const srcFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
    const destFolder = DriveApp.getFolderById(DEST_FOLDER_ID);
    const files = srcFolder.getFiles();

    while (files.hasNext()) {
      // 【追加1】GASの6分制限が近づいたら処理を安全に中断
      if (Date.now() - startTime > MAX_EXECUTION_TIME) {
        Logger.log("【警告】GASの6分制限が近づいたため、処理を安全に中断します。残りのファイルは次回実行時に処理されます。");
        break;
      }

      const file = files.next();
      const mimeType = file.getMimeType();

      if (!mimeType.includes('pdf') && !mimeType.includes('image')) continue;

      Logger.log(`処理開始: ${file.getName()}`);
      await processFile(file, mimeType, destFolder);

      // API制限回避の待機（60秒）
      if (files.hasNext()) {
        Logger.log(`次のファイル処理まで60秒待機します...`);
        Utilities.sleep(60000);
      }
    }
  } catch (e) {
    Logger.log(`【エラー発生】: ` + e.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * 1ファイル分の処理（AI解析 → 向き補正 → リネーム＆移動）
 */
async function processFile(file, mimeType, destFolder) {
  const bytes = file.getBlob().getBytes();
  const isPdf = mimeType === MimeType.PDF;
  const isConvertibleImage = mimeType === MimeType.JPEG || mimeType === MimeType.PNG;

  // PDFはページ数を取得してGeminiに伝える（ページごとの向き判定のため）
  let pdfDoc = null;
  let pageCount = 1;
  if (isPdf) {
    try {
      const { PDFDocument } = loadPdfLib();
      pdfDoc = await PDFDocument.load(new Uint8Array(bytes), { ignoreEncryption: true });
      pageCount = pdfDoc.getPageCount();
    } catch (err) {
      Logger.log(`PDFの読み込みに失敗したため、向き補正は行いません: ${err.toString()}`);
      pdfDoc = null;
    }
  }

  // AI処理の実行（ファイル名＋各ページの回転角度を取得）
  const result = getNewFilenameFromGemini(Utilities.base64Encode(bytes), mimeType, pageCount);
  if (!result) {
    Logger.log(`【失敗】ファイル名を取得できませんでした`);
    return;
  }

  const rotations = result.rotations;
  const needsRotation = rotations.some(r => r !== 0);
  Logger.log(`回転判定: [${rotations.join(', ')}]`);

  // ① JPEG/PNG → 向きを補正してPDFに変換（新規作成＋元ファイルはゴミ箱へ）
  if (isConvertibleImage) {
    const newName = withExtension(result.baseName, 'pdf');
    const pdfBlob = await imageToRotatedPdf(bytes, mimeType, rotations[0], newName);
    const newFile = destFolder.createFile(pdfBlob);
    file.setTrashed(true);
    Logger.log(`【成功】PDF変換${needsRotation ? '＋向き補正' : ''}＆移動完了: ${newName}`);
    Logger.log(`保存先リンク: ${newFile.getUrl()}`);
    return;
  }

  // ② PDFで回転が必要 → 回転済みPDFを新規作成＋元ファイルはゴミ箱へ
  if (isPdf && pdfDoc && needsRotation) {
    const newName = withExtension(result.baseName, 'pdf');
    const pdfBlob = await rotatePdf(pdfDoc, rotations, newName);
    const newFile = destFolder.createFile(pdfBlob);
    file.setTrashed(true);
    Logger.log(`【成功】向き補正＆リネーム＆移動完了: ${newName}`);
    Logger.log(`保存先リンク: ${newFile.getUrl()}`);
    return;
  }

  // ③ 回転不要（またはHEIC等の変換非対応画像）→ 従来どおりリネーム＆移動のみ
  const ext = isPdf ? 'pdf' : (getExtension(file.getName()) || 'jpg');
  const newName = withExtension(result.baseName, ext);
  if (!isPdf && needsRotation) {
    Logger.log(`※この画像形式（${mimeType}）は向き補正に対応していないため、回転せずに移動します。`);
  }
  file.setName(newName);
  file.moveTo(destFolder);
  Logger.log(`【成功】リネーム＆移動完了: ${newName}`);
  Logger.log(`保存先リンク: ${file.getUrl()}`);
}

/**
 * Geminiでファイル名と各ページの回転角度を取得
 * 戻り値: { baseName: '20260929_領収書_〇〇不動産', rotations: [0, 90, ...] } / 失敗時 null
 */
function getNewFilenameFromGemini(base64Data, mimeType, pageCount) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    Logger.log("エラー: APIキーがスクリプトプロパティに設定されていません。");
    return null;
  }

  const url = `https://generativelanguage.googleapis.com/v1/models/gemini-3.8-flash:generateContent?key=${apiKey}`;

  const prompt =
    "あなたは書類のスキャンデータを自動で整理するアシスタントです。\n" +
    "添付データ（全" + pageCount + "ページ）について、次の2点を判定してください。\n" +
    "1. 内容（日付、書類種別、相手先名、金額等）を読み取り、『YYYYMMDD_書類種別_相手先名.pdf』形式の適切なファイル名\n" +
    "2. 各ページの文字が正しく読める向き（正立）にするために、時計回りに何度回転させる必要があるか（0/90/180/270のいずれか）。" +
    "すでに正しい向きのページは0。横向きの書類（表や横長の帳票）が正しく横向きで読めている場合も0。\n" +
    "出力は次のJSONのみとし、余計なテキストや解説、コードブロック記号は一切含めないでください。\n" +
    '{"filename":"20260929_領収書_〇〇不動産.pdf","rotations":[' +
    Array(pageCount).fill(0).join(',') + ']}\n' +
    "※rotationsは1ページ目から順に全" + pageCount + "ページ分を出力してください。";

  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        {
          "inline_data": {
            "mime_type": mimeType,
            "data": base64Data
          }
        }
      ]
    }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  // 【追加2】エラー時（503や429）に、最大3回までリトライ（再挑戦）する仕組み
  let maxRetries = 3;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseText = response.getContentText();
      const responseCode = response.getResponseCode();

      // 成功した場合（HTTPステータス200番台）
      if (responseCode >= 200 && responseCode < 300) {
        const json = JSON.parse(responseText);
        const parts = (json.candidates && json.candidates[0].content && json.candidates[0].content.parts) || [];
        const text = parts.map(p => p.text || '').join('').trim();
        const parsed = parseGeminiResult(text, pageCount);
        if (parsed) return parsed;

        Logger.log(`APIレスポンスの形式が想定外です: ${responseText}`);
        return null; // 形式エラーの場合は再試行せず終了
      }
      // エラーが発生した場合（429や503など）
      else {
        Logger.log(`【エラー】API呼び出し失敗 (ステータス: ${responseCode}) - 試行回数: ${attempt}/${maxRetries}`);
        Logger.log(`詳細: ${responseText}`);

        // もしこれが最後の挑戦だったら、諦めてnullを返す
        if (attempt === maxRetries) {
          Logger.log("規定の再試行回数に達したため、処理を断念します。");
          return null;
        }

        // 次の再試行まで少し待つ（1回目は20秒、2回目は40秒...と増やす）
        const waitTime = attempt * 20000;
        Logger.log(`${waitTime/1000}秒後に再試行します...`);
        Utilities.sleep(waitTime);
      }
    } catch (err) {
      Logger.log(`通信エラーまたはJSONパースエラー: ${err.toString()}`);
      return null;
    }
  }
}

/**
 * Geminiの出力テキスト（JSON）を解析
 */
function parseGeminiResult(text, pageCount) {
  const match = text.replace(/```(json)?/g, '').match(/\{[\s\S]*\}/);
  if (!match) return null;

  let obj;
  try {
    obj = JSON.parse(match[0]);
  } catch (e) {
    return null;
  }
  if (!obj.filename) return null;

  // ファイル名の整形（禁止文字除去・拡張子除去）
  let baseName = String(obj.filename)
    .replace(/`/g, '').replace(/\n/g, '').replace(/[\/\?<>\\:\*\|"]/g, '')
    .trim()
    .replace(/\.(pdf|jpe?g|png|heic|tiff?)$/i, '');
  if (!baseName) return null;

  // 回転角度の正規化（0/90/180/270以外は丸める。不足ページは0扱い）
  const raw = Array.isArray(obj.rotations) ? obj.rotations : [];
  const rotations = [];
  for (let i = 0; i < pageCount; i++) {
    const r = Number(raw[i]) || 0;
    rotations.push((((Math.round(r / 90) * 90) % 360) + 360) % 360);
  }

  return { baseName: baseName, rotations: rotations };
}

/**
 * PDFの各ページを指定角度（時計回り）で回転させたBlobを返す
 */
async function rotatePdf(pdfDoc, rotations, fileName) {
  const { degrees } = loadPdfLib();
  pdfDoc.getPages().forEach((page, i) => {
    const r = rotations[i] || 0;
    if (r === 0) return;
    const current = page.getRotation().angle || 0;
    page.setRotation(degrees((current + r) % 360));
  });
  const out = await pdfDoc.save();
  return Utilities.newBlob([...out], MimeType.PDF, fileName);
}

/**
 * JPEG/PNG画像を、向きを補正した1ページのPDFに変換したBlobを返す
 */
async function imageToRotatedPdf(bytes, mimeType, rotation, fileName) {
  const { PDFDocument, degrees } = loadPdfLib();
  const doc = await PDFDocument.create();
  const u8 = new Uint8Array(bytes);
  const img = mimeType === MimeType.PNG ? await doc.embedPng(u8) : await doc.embedJpg(u8);

  const scale = PAGE_LONG_SIDE_PT / Math.max(img.width, img.height);
  const w = img.width * scale;
  const h = img.height * scale;

  const page = doc.addPage([w, h]);
  page.drawImage(img, { x: 0, y: 0, width: w, height: h });
  if (rotation) page.setRotation(degrees(rotation));

  const out = await doc.save();
  return Utilities.newBlob([...out], MimeType.PDF, fileName);
}

/**
 * pdf-libの読み込み（1実行につき1回だけ取得）
 */
function loadPdfLib() {
  if (typeof PDFLib !== 'undefined') return PDFLib;
  // pdf-libをGASで動かすためのsetTimeout代替
  const setTimeout = function (f, t) { Utilities.sleep(t || 0); return f(); };
  eval(UrlFetchApp.fetch(PDF_LIB_URL).getContentText());
  return PDFLib;
}

function getExtension(name) {
  const m = String(name).match(/\.([A-Za-z0-9]+)$/);
  return m ? m[1].toLowerCase() : '';
}

function withExtension(baseName, ext) {
  return `${baseName}.${ext}`;
}

/**
 * 【起動方法①】定期実行トリガーの設定（最初に1回だけ手動実行）
 */
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkForNewFiles')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkForNewFiles').timeBased().everyMinutes(10).create();
}

/**
 * 【起動方法②】WebアプリURLを開くと即時実行（スマホのホーム画面に置くと便利）
 */
function doGet() {
  ScriptApp.newTrigger('runOnce').timeBased().after(1000).create();
  return HtmlService.createHtmlOutput('<p style="font-size:20px">処理を開始しました。1〜数分後に「処理済みデータ」をご確認ください。</p>');
}

async function runOnce() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runOnce')
    .forEach(t => ScriptApp.deleteTrigger(t));
  await checkForNewFiles();
}
