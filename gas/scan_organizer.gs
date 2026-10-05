/**
 * スキャン整理（ファイル名変更＆移動＋向き補正）
 *
 * 【処理①】checkForNewFiles … 未処理データ → AIでファイル名判定・向き補正 → 処理済みデータへ移動（5分ごと）
 * 【処理②】fixOrientation  … 処理済みデータ内の既存PDFの向きを上書き修正（10分ごと）
 *          ※処理①で向き判定済みのファイルには目印を付けるため、処理②で再判定はしない
 *
 * 【事前準備】
 * ・「サービス」→「Drive API」（v3）を追加
 * ・スクリプトプロパティに GEMINI_API_KEY を設定
 * ・setupTrigger() を1回手動実行（処理①②のトリガーをまとめて設定）
 */

const SOURCE_FOLDER_ID = '1Oln0vKLfmNInPyHVMPnR22tEtp9KXozB'; // マイドライブ/scan/未処理データ
const DEST_FOLDER_ID   = '1bTo9kY4ZJQCtMeURLJZEHOTIHoyfWJJN'; // マイドライブ/scan/処理済みデータ

// PDFの回転・画像→PDF変換に使用するライブラリ（pdf-lib）
const PDF_LIB_URL = 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js';

// 画像をPDF化する際のページサイズ（長辺をA4相当 842pt に合わせる）
const PAGE_LONG_SIDE_PT = 842;

// 使用するGeminiモデル（メイン）
const PRIMARY_MODEL = 'gemini-3.8-flash';
// メインが混雑（503等）の場合に切り替える代替モデルの数（利用可能なflash系モデルから自動選定）
const MAX_FALLBACK_MODELS = 2;

// 向き補正の設定
// ENABLE_ROTATION = false にすると、向きは変えずにファイル名の変更＆移動のみ行う
const ENABLE_ROTATION = true;
// この容量を超えるPDFは向き補正を行わない（GASのメモリ・時間制限対策）
const ROTATE_MAX_BYTES = 5 * 1024 * 1024;   // 5MB

// この容量を超えるファイルは、Gemini File API（アップロード方式）で送信する
// （直接送信はBase64化で約1.33倍に膨らみ、20MBの上限や混雑時の失敗に掛かりやすいため）
const INLINE_MAX_BYTES = 8 * 1024 * 1024;   // 8MB
// GAS（UrlFetchApp）で送信できる上限（50MB）。これを超えるファイルは処理しない
const UPLOAD_MAX_BYTES = 48 * 1024 * 1024;  // 48MB

// 【処理②】既存PDFの向き補正の設定
const ORIENT_INCLUDE_SUBFOLDERS = false;  // true にすると処理済みデータのサブフォルダ内も対象
const ORIENT_MAX_FAILS = 3;               // 失敗がこの回数に達したファイルは以後スキップ
// 判定済みの目印（Driveのファイルに付与する非表示プロパティ）
const ORIENT_PROP_KEY = 'orientFix';      // 値: done（判定済み）/ large（容量超過）/ error（失敗上限）
const ORIENT_FAIL_KEY = 'orientFixFails';

// 実行全体の締切（GASの6分制限に掛からないよう、再試行の待機もこの時刻までに収める）
let RUN_DEADLINE = 0;


/* =====================================================================
 * 【処理①】未処理データの整理（ファイル名変更＆移動＋向き補正）
 * ===================================================================== */

async function checkForNewFiles() {
  const startTime = Date.now(); // 実行開始時間を記録
  const START_CUTOFF = 2 * 60 * 1000; // 新しいファイルの処理開始は実行開始から2分以内に限る（1ファイル最大3〜4分かかるため）
  RUN_DEADLINE = startTime + 5.5 * 60 * 1000;  // 再試行は開始から5.5分以内に収める

  // 二重起動防止（前回の処理が実行中なら今回はスキップ）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log("前回の処理が実行中のため、今回はスキップします。");
    return;
  }

  try {
    const srcFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
    const destFolder = DriveApp.getFolderById(DEST_FOLDER_ID);
    const props = PropertiesService.getScriptProperties();

    // 対象ファイルを一覧化し、失敗回数の少ない順→古い順に並べる
    // （失敗し続けるファイルが先頭に居座り、他のファイルが処理されなくなるのを防ぐ）
    const targets = [];
    const it = srcFolder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      const t = f.getMimeType();
      if (!t.includes('pdf') && !t.includes('image')) continue;
      targets.push({ file: f, mimeType: t, fails: Number(props.getProperty('fail_' + f.getId()) || 0), created: f.getDateCreated().getTime() });
    }
    targets.sort((a, b) => (a.fails - b.fails) || (a.created - b.created));
    Logger.log(`未処理データ: ${targets.length}件`);

    for (let i = 0; i < targets.length; i++) {
      // GASの6分制限が近づいたら処理を安全に中断
      if (Date.now() - startTime > START_CUTOFF) {
        Logger.log("【警告】GASの6分制限に備え、今回の処理はここまでとします。残りのファイルは次回実行時に処理されます。");
        break;
      }

      const { file, mimeType, fails } = targets[i];
      const failKey = 'fail_' + file.getId();
      Logger.log(`処理開始: ${file.getName()}${fails ? `（過去の失敗: ${fails}回）` : ''}`);

      // 処理前に失敗回数を加算（6分超過で強制終了された場合も失敗として数えるため）。成功したら削除
      props.setProperty(failKey, String(fails + 1));
      let ok = false;
      try {
        ok = await processFile(file, mimeType, destFolder);
      } catch (err) {
        Logger.log(`【エラー】${file.getName()} の処理中にエラーが発生しました: ${err.toString()}`);
      }
      if (ok) props.deleteProperty(failKey);

      // API制限回避の待機（15秒）※待機後に処理開始期限を過ぎる場合は待たずに終了
      if (i < targets.length - 1) {
        if (Date.now() - startTime + 15000 > START_CUTOFF) {
          Logger.log("今回の処理はここまでとします。残りのファイルは次回実行時に処理されます。");
          break;
        }
        Logger.log(`次のファイル処理まで15秒待機します...`);
        Utilities.sleep(15000);
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
  // ※大容量ファイルはバイト配列に展開するとGASのメモリ不足になるため、Blobのまま扱い、必要な場合のみ展開する
  const blob = file.getBlob();
  const size = file.getSize();
  const isPdf = mimeType === MimeType.PDF;
  const isConvertibleImage = ENABLE_ROTATION && (mimeType === MimeType.JPEG || mimeType === MimeType.PNG);

  // 向き補正を行うか判定（無効設定、または大容量PDFの場合はファイル名の変更のみ）
  const rotate = ENABLE_ROTATION && !(isPdf && size > ROTATE_MAX_BYTES);
  const bytes = rotate ? blob.getBytes() : null;  // 向き補正時のみ展開
  if (ENABLE_ROTATION && !rotate) {
    Logger.log(`容量が大きいため（${(size / 1024 / 1024).toFixed(1)}MB）、向き補正は行わずファイル名の変更のみ行います。`);
  }

  // PDFはページ数を取得してGeminiに伝える（ページごとの向き判定のため）
  // ※向き補正を行わない場合は pageCount = 0（ファイル名のみ判定）
  let pdfDoc = null;
  let pageCount = rotate ? 1 : 0;
  if (isPdf && rotate) {
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
  const result = getNewFilenameFromGemini(blob, size, mimeType, pageCount, file.getName());
  if (!result) {
    Logger.log(`【失敗】ファイル名を取得できませんでした`);
    return false;
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
    markOrientation(newFile.getId(), { [ORIENT_PROP_KEY]: 'done' });
    Logger.log(`【成功】PDF変換${needsRotation ? '＋向き補正' : ''}＆移動完了: ${newName}`);
    Logger.log(`保存先リンク: ${newFile.getUrl()}`);
    return true;
  }

  // ② PDFで回転が必要 → 回転済みPDFを新規作成＋元ファイルはゴミ箱へ
  if (isPdf && pdfDoc && needsRotation) {
    const newName = withExtension(result.baseName, 'pdf');
    const pdfBlob = await rotatePdf(pdfDoc, rotations, newName);
    const newFile = destFolder.createFile(pdfBlob);
    file.setTrashed(true);
    markOrientation(newFile.getId(), { [ORIENT_PROP_KEY]: 'done' });
    Logger.log(`【成功】向き補正＆リネーム＆移動完了: ${newName}`);
    Logger.log(`保存先リンク: ${newFile.getUrl()}`);
    return true;
  }

  // ③ 回転不要（またはHEIC等の変換非対応画像）→ 従来どおりリネーム＆移動のみ
  const ext = isPdf ? 'pdf' : (getExtension(file.getName()) || 'jpg');
  const newName = withExtension(result.baseName, ext);
  if (!isPdf && needsRotation) {
    Logger.log(`※この画像形式（${mimeType}）は向き補正に対応していないため、回転せずに移動します。`);
  }
  file.setName(newName);
  file.moveTo(destFolder);
  if (isPdf && pdfDoc) markOrientation(file.getId(), { [ORIENT_PROP_KEY]: 'done' });  // 向き判定済み（回転不要）
  Logger.log(`【成功】リネーム＆移動完了: ${newName}`);
  Logger.log(`保存先リンク: ${file.getUrl()}`);
  return true;
}

/**
 * Geminiでファイル名と各ページの回転角度を取得
 * 戻り値: { baseName: '20260929_領収書_〇〇不動産', rotations: [0, 90, ...] } / 失敗時 null
 */
function getNewFilenameFromGemini(blob, size, mimeType, pageCount, displayName) {
  const apiKey = getApiKey();
  if (!apiKey) return null;

  const sizeMB = (size / 1024 / 1024).toFixed(1);
  if (size > UPLOAD_MAX_BYTES) {
    Logger.log(`【スキップ】ファイル容量が大きすぎるため処理できません（${sizeMB}MB）。スキャン設定（解像度・カラー）を下げて取り込み直してください。`);
    return null;
  }

  // 容量に応じて送信方法を切り替え（小：直接送信／大：File APIでアップロードして参照）
  let filePart;
  let uploaded = null;
  if (size <= INLINE_MAX_BYTES) {
    filePart = { "inline_data": { "mime_type": mimeType, "data": Utilities.base64Encode(blob.getBytes()) } };
  } else {
    Logger.log(`ファイル容量が大きいため（${sizeMB}MB）、Gemini File APIでアップロードします...`);
    uploaded = uploadToGeminiFileApi(blob, size, mimeType, displayName, apiKey);
    if (!uploaded) return null;
    filePart = { "file_data": { "mime_type": mimeType, "file_uri": uploaded.uri } };
  }

  const prompt = pageCount === 0
    // 向き補正なし：ファイル名のみ判定
    ? "あなたは書類のスキャンデータを自動で整理するアシスタントです。\n" +
      "添付データの内容（日付、書類種別、相手先名、金額等）を読み取り、『YYYYMMDD_書類種別_相手先名.pdf』形式の適切なファイル名を判定してください。\n" +
      "出力は次のJSONのみとし、余計なテキストや解説、コードブロック記号は一切含めないでください。\n" +
      '{"filename":"20260929_領収書_〇〇不動産.pdf"}'
    // 向き補正あり：ファイル名＋各ページの回転角度を判定
    : "あなたは書類のスキャンデータを自動で整理するアシスタントです。\n" +
      "添付データ（全" + pageCount + "ページ）について、次の2点を判定してください。\n" +
      "1. 内容（日付、書類種別、相手先名、金額等）を読み取り、『YYYYMMDD_書類種別_相手先名.pdf』形式の適切なファイル名\n" +
      "2. 各ページの文字が正しく読める向き（正立）にするために、時計回りに何度回転させる必要があるか（0/90/180/270のいずれか）。" +
      "すでに正しい向きのページは0。横向きの書類（表や横長の帳票）が正しく横向きで読めている場合も0。\n" +
      "出力は次のJSONのみとし、余計なテキストや解説、コードブロック記号は一切含めないでください。\n" +
      '{"filename":"20260929_領収書_〇〇不動産.pdf","rotations":[' +
      Array(pageCount).fill(0).join(',') + ']}\n' +
      "※rotationsは1ページ目から順に全" + pageCount + "ページ分を出力してください。";

  try {
    const obj = callGemini(apiKey, prompt, filePart);
    if (!obj || !obj.filename) return null;

    // ファイル名の整形（禁止文字除去・拡張子除去）
    const baseName = String(obj.filename)
      .replace(/`/g, '').replace(/\n/g, '').replace(/[\/\?<>\\:\*\|"]/g, '')
      .trim()
      .replace(/\.(pdf|jpe?g|png|heic|tiff?)$/i, '');
    if (!baseName) return null;

    return { baseName: baseName, rotations: normalizeRotations(obj.rotations, pageCount) };
  } finally {
    if (uploaded) deleteGeminiFile(uploaded.name, apiKey);
  }
}


/* =====================================================================
 * 【処理②】処理済みデータ内の既存PDFの向き補正（元ファイルを上書き修正）
 *  ・回転はページの /Rotate 属性の変更のみ（画像の再圧縮なし）→ 容量はほぼ増えない
 *  ・ファイルID・URL・共有設定・ファイル名はそのまま
 * ===================================================================== */

async function fixOrientation() {
  const startTime = Date.now();
  const START_CUTOFF = 4 * 60 * 1000;  // 新しいファイルの処理開始は4分以内
  RUN_DEADLINE = startTime + 5.5 * 60 * 1000;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('他の処理が実行中のため、今回はスキップします。');
    return;
  }

  try {
    const apiKey = getApiKey();
    if (!apiKey) return;

    const targets = listOrientationTargets(DEST_FOLDER_ID);
    Logger.log(`向き補正の対象（未判定）: ${targets.length}件`);
    let fixed = 0, unchanged = 0;

    for (let i = 0; i < targets.length; i++) {
      if (Date.now() - startTime > START_CUTOFF) {
        Logger.log('実行時間の上限が近いため、残りは次回実行時に処理します。');
        break;
      }
      const f = targets[i];
      const size = Number(f.size || 0);

      if (size > ROTATE_MAX_BYTES) {
        Logger.log(`【スキップ】容量超過（${(size / 1024 / 1024).toFixed(1)}MB）: ${f.name}`);
        markOrientation(f.id, { [ORIENT_PROP_KEY]: 'large' });
        continue;
      }

      Logger.log(`判定開始: ${f.name}`);
      try {
        const result = await fixOrientationOfFile(f, apiKey);
        if (result === 'fixed') fixed++;
        else if (result === 'unchanged') unchanged++;
        else countOrientationFail(f);
      } catch (err) {
        Logger.log(`【エラー】${f.name}: ${err.toString()}`);
        countOrientationFail(f);
      }

      if (i < targets.length - 1) Utilities.sleep(5000);  // API制限回避
    }
    Logger.log(`完了: 補正 ${fixed}件 ／ 補正不要 ${unchanged}件`);
  } catch (e) {
    Logger.log(`【エラー発生】: ${e.toString()}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 1ファイル分の向き補正
 * 戻り値: 'fixed'（回転して上書き）/ 'unchanged'（回転不要）/ null（失敗）
 */
async function fixOrientationOfFile(f, apiKey) {
  const bytes = DriveApp.getFileById(f.id).getBlob().getBytes();

  const { PDFDocument } = loadPdfLib();
  let pdfDoc;
  try {
    pdfDoc = await PDFDocument.load(new Uint8Array(bytes), { ignoreEncryption: true, updateMetadata: false });
  } catch (err) {
    Logger.log(`PDFの読み込みに失敗しました: ${err.toString()}`);
    return null;
  }
  const pageCount = pdfDoc.getPageCount();

  const prompt =
    'あなたは書類のスキャンデータの向きを判定するアシスタントです。\n' +
    '添付PDF（全' + pageCount + 'ページ）の各ページについて、文字が正しく読める向き（正立）にするために' +
    '時計回りに何度回転させる必要があるかを判定してください（0/90/180/270のいずれか）。\n' +
    '・縦向きの書類が横倒し・逆さまになっている場合は、正しく縦向きで読める角度\n' +
    '・横向きの書類（表や横長の帳票）が正しく横向きで読めている場合は0\n' +
    '・すでに正しい向きのページ、白紙のページは0\n' +
    '出力は次のJSONのみとし、余計なテキストや解説、コードブロック記号は一切含めないでください。\n' +
    '{"rotations":[' + Array(pageCount).fill(0).join(',') + ']}\n' +
    '※rotationsは1ページ目から順に全' + pageCount + 'ページ分を出力してください。';
  const filePart = { inline_data: { mime_type: 'application/pdf', data: Utilities.base64Encode(bytes) } };

  const obj = callGemini(apiKey, prompt, filePart);
  if (!obj || !Array.isArray(obj.rotations)) return null;
  const rotations = normalizeRotations(obj.rotations, pageCount);
  Logger.log(`回転判定: [${rotations.join(', ')}]`);

  // 回転不要 → ファイルには触れず、判定済みの目印のみ付与
  if (!rotations.some(r => r !== 0)) {
    markOrientation(f.id, { [ORIENT_PROP_KEY]: 'done', [ORIENT_FAIL_KEY]: null });
    Logger.log('補正不要');
    return 'unchanged';
  }

  // 元ファイルを上書き（ファイルID・URL・共有設定・ファイル名は維持）
  const newBlob = await rotatePdf(pdfDoc, rotations, f.name);
  Drive.Files.update(
    { appProperties: { [ORIENT_PROP_KEY]: 'done', [ORIENT_FAIL_KEY]: null } },
    f.id,
    newBlob
  );
  Logger.log(`【補正完了】${f.name}（${(bytes.length / 1024).toFixed(0)}KB → ${(newBlob.getBytes().length / 1024).toFixed(0)}KB）`);
  return 'fixed';
}

/**
 * 未判定のPDFを一覧化（判定済み・容量超過・失敗上限のものは除外）
 */
function listOrientationTargets(folderId) {
  const results = [];
  const folders = [folderId];
  while (folders.length) {
    const id = folders.shift();
    let pageToken;
    do {
      const res = Drive.Files.list({
        q: `'${id}' in parents and trashed = false and (mimeType = 'application/pdf' or mimeType = 'application/vnd.google-apps.folder')`,
        fields: 'nextPageToken, files(id, name, mimeType, size, createdTime, appProperties)',
        pageSize: 1000,
        pageToken: pageToken
      });
      (res.files || []).forEach(f => {
        if (f.mimeType === 'application/vnd.google-apps.folder') {
          if (ORIENT_INCLUDE_SUBFOLDERS) folders.push(f.id);
          return;
        }
        if ((f.appProperties || {})[ORIENT_PROP_KEY]) return;
        results.push(f);
      });
      pageToken = res.nextPageToken;
    } while (pageToken);
  }
  results.sort((a, b) => new Date(a.createdTime) - new Date(b.createdTime));  // 古い順
  return results;
}

/**
 * 向き判定の目印（appProperties）を付与 ※ファイル本体は変更しない
 */
function markOrientation(fileId, props) {
  try {
    Drive.Files.update({ appProperties: props }, fileId);
  } catch (e) {
    Logger.log(`目印の付与に失敗しました（処理には影響なし）: ${e.toString()}`);
  }
}

function countOrientationFail(f) {
  const fails = Number((f.appProperties || {})[ORIENT_FAIL_KEY] || 0) + 1;
  if (fails >= ORIENT_MAX_FAILS) {
    Logger.log(`失敗が${fails}回に達したため、以後この判定はスキップします: ${f.name}`);
    markOrientation(f.id, { [ORIENT_PROP_KEY]: 'error' });
  } else {
    markOrientation(f.id, { [ORIENT_FAIL_KEY]: String(fails) });
  }
}


/* =====================================================================
 * 共通処理（Gemini・PDF）
 * ===================================================================== */

function getApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) Logger.log("エラー: APIキーがスクリプトプロパティに設定されていません。");
  return apiKey;
}

/**
 * Geminiへの問い合わせ（モデル切替＋リトライ）
 * 戻り値: 出力JSONを解析したオブジェクト / 失敗時 null
 */
function callGemini(apiKey, prompt, filePart) {
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({ "contents": [{ "parts": [{ "text": prompt }, filePart] }] }),
    "muteHttpExceptions": true
  };

  // エラー時（503や429）に、モデルを切り替えながら最大3回までリトライ
  const models = getModelCandidates(apiKey);
  const maxRetries = 3;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const model = models[(attempt - 1) % models.length];
    const nextModel = models[attempt % models.length];
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    try {
      const response = UrlFetchApp.fetch(url, options);
      const responseText = response.getContentText();
      const responseCode = response.getResponseCode();

      // 成功した場合（HTTPステータス200番台）
      if (responseCode >= 200 && responseCode < 300) {
        const json = JSON.parse(responseText);
        const parts = (json.candidates && json.candidates[0].content && json.candidates[0].content.parts) || [];
        const text = parts.map(p => p.text || '').join('').replace(/```(json)?/g, '');
        const match = text.match(/\{[\s\S]*\}/);
        if (match) {
          Logger.log(`使用モデル: ${model}`);
          return JSON.parse(match[0]);
        }
        Logger.log(`APIレスポンスの形式が想定外です: ${responseText}`);
        return null; // 形式エラーの場合は再試行せず終了
      }

      // エラーが発生した場合（429や503など）
      Logger.log(`【エラー】API呼び出し失敗 (ステータス: ${responseCode} / モデル: ${model}) - 試行回数: ${attempt}/${maxRetries}`);
      Logger.log(`詳細: ${responseText}`);
      if (attempt === maxRetries) {
        Logger.log("規定の再試行回数に達したため、処理を断念します。");
        return null;
      }

      // 次の再試行まで少し待つ（別モデルに切り替える場合は5秒、同じモデルの場合は20秒→40秒と増やす）
      const waitTime = nextModel !== model ? 5000 : attempt * 20000;
      // 待機＋再呼び出し（約1分）で6分制限を超えそうなら、今回は見送り（次回実行時に再処理）
      if (RUN_DEADLINE && Date.now() + waitTime + 60000 > RUN_DEADLINE) {
        Logger.log("実行時間の上限が近いため再試行を見送ります。このファイルは次回実行時に再処理されます。");
        return null;
      }
      Logger.log(`${waitTime / 1000}秒後に再試行します（モデル: ${nextModel}）...`);
      Utilities.sleep(waitTime);
    } catch (err) {
      Logger.log(`通信エラーまたはJSONパースエラー: ${err.toString()}`);
      return null;
    }
  }
  return null;
}

/**
 * 回転角度の正規化（0/90/180/270以外は丸める。不足ページは0扱い）
 */
function normalizeRotations(raw, pageCount) {
  const list = Array.isArray(raw) ? raw : [];
  const rotations = [];
  for (let i = 0; i < pageCount; i++) {
    const r = Number(list[i]) || 0;
    rotations.push((((Math.round(r / 90) * 90) % 360) + 360) % 360);
  }
  return rotations;
}

/**
 * Gemini File APIへファイルをアップロード（再開可能アップロード）
 * 戻り値: { name: 'files/xxxx', uri: 'https://...' } / 失敗時 null
 */
function uploadToGeminiFileApi(blob, size, mimeType, displayName, apiKey) {
  try {
    // ① アップロード開始（アップロード先URLを取得）
    const startRes = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/upload/v1beta/files?key=${apiKey}`, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'X-Goog-Upload-Protocol': 'resumable',
        'X-Goog-Upload-Command': 'start',
        'X-Goog-Upload-Header-Content-Length': String(size),
        'X-Goog-Upload-Header-Content-Type': mimeType
      },
      payload: JSON.stringify({ file: { display_name: displayName } }),
      muteHttpExceptions: true
    });
    const headers = startRes.getAllHeaders();
    const uploadUrlKey = Object.keys(headers).find(k => k.toLowerCase() === 'x-goog-upload-url');
    if (!uploadUrlKey) {
      Logger.log(`File APIのアップロード開始に失敗しました (ステータス: ${startRes.getResponseCode()}): ${startRes.getContentText()}`);
      return null;
    }

    // ② ファイル本体を送信
    const upRes = UrlFetchApp.fetch(headers[uploadUrlKey], {
      method: 'post',
      contentType: mimeType,
      headers: { 'X-Goog-Upload-Offset': '0', 'X-Goog-Upload-Command': 'upload, finalize' },
      payload: blob,  // Blobのまま送信（バイト配列に展開しない）
      muteHttpExceptions: true
    });
    if (upRes.getResponseCode() !== 200) {
      Logger.log(`File APIへのアップロードに失敗しました (ステータス: ${upRes.getResponseCode()}): ${upRes.getContentText()}`);
      return null;
    }
    let info = JSON.parse(upRes.getContentText()).file;

    // ③ Gemini側の取り込み完了（ACTIVE）を待つ（最大60秒）
    for (let i = 0; i < 12 && info.state === 'PROCESSING'; i++) {
      Utilities.sleep(5000);
      const st = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/${info.name}?key=${apiKey}`, { muteHttpExceptions: true });
      info = JSON.parse(st.getContentText());
    }
    if (info.state && info.state !== 'ACTIVE') {
      Logger.log(`File APIでの取り込みが完了しませんでした (状態: ${info.state})`);
      deleteGeminiFile(info.name, apiKey);
      return null;
    }

    Logger.log(`File APIへのアップロード完了: ${info.name}`);
    return { name: info.name, uri: info.uri };
  } catch (e) {
    Logger.log(`File APIへのアップロード中にエラーが発生しました: ${e.toString()}`);
    return null;
  }
}

/**
 * File APIにアップロードしたファイルを削除（※削除しなくても48時間で自動削除）
 */
function deleteGeminiFile(name, apiKey) {
  try {
    UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/${name}?key=${apiKey}`, { method: 'delete', muteHttpExceptions: true });
  } catch (e) {
    // 削除失敗は処理に影響しないため無視
  }
}

/**
 * 使用するモデルの候補リストを取得（メイン＋代替モデル）
 * 代替モデルは、このAPIキーで利用可能なflash系モデルから新しい順に自動選定（6時間キャッシュ）
 */
function getModelCandidates(apiKey) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'gemini_models_' + PRIMARY_MODEL;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  let list = [PRIMARY_MODEL];
  try {
    const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models?pageSize=1000&key=${apiKey}`, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const version = n => { const m = n.match(/gemini-(\d+(?:\.\d+)?)/); return m ? parseFloat(m[1]) : 0; };
      const fallbacks = (JSON.parse(res.getContentText()).models || [])
        .filter(m => (m.supportedGenerationMethods || []).includes('generateContent'))
        .map(m => m.name.replace('models/', ''))
        .filter(n => /flash/.test(n) && !/(image|tts|audio|live|embedding|exp)/.test(n) && n !== PRIMARY_MODEL)
        .sort((a, b) =>
          (/preview/.test(a) - /preview/.test(b)) ||   // 正式版を優先
          (version(b) - version(a)) ||                 // 新しい世代を優先
          (/lite/.test(a) - /lite/.test(b)));          // 通常版をlite版より優先
      list = list.concat([...new Set(fallbacks)].slice(0, MAX_FALLBACK_MODELS));
    } else {
      Logger.log(`モデル一覧の取得に失敗したため、メインモデルのみで実行します（ステータス: ${res.getResponseCode()}）`);
    }
  } catch (e) {
    Logger.log(`モデル一覧の取得に失敗したため、メインモデルのみで実行します: ${e.toString()}`);
  }

  Logger.log(`使用モデル候補: ${list.join(' → ')}`);
  cache.put(cacheKey, JSON.stringify(list), 6 * 60 * 60);
  return list;
}

/**
 * PDFの各ページを指定角度（時計回り）で回転させたBlobを返す
 * ※ページの /Rotate 属性の変更のみ（画像データはそのまま）
 */
async function rotatePdf(pdfDoc, rotations, fileName) {
  const { degrees } = loadPdfLib();
  pdfDoc.getPages().forEach((page, i) => {
    const r = rotations[i] || 0;
    if (r === 0) return;
    const current = page.getRotation().angle || 0;
    page.setRotation(degrees((current + r) % 360));
  });
  const out = await pdfDoc.save({ useObjectStreams: true });
  return Utilities.newBlob(Array.from(out), MimeType.PDF, fileName);
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
  return Utilities.newBlob(Array.from(out), MimeType.PDF, fileName);
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


/* =====================================================================
 * 起動・補助
 * ===================================================================== */

/**
 * 【起動方法①】定期実行トリガーの設定（最初に1回だけ手動実行）
 * 処理①：5分ごと／処理②：10分ごと
 */
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => ['checkForNewFiles', 'fixOrientation', 'ori_fixOrientation'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkForNewFiles').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('fixOrientation').timeBased().everyMinutes(10).create();
}

/**
 * 【起動方法②】WebアプリURLを開くと即時実行（スマホのホーム画面に置くと便利）
 */
function doGet() {
  ScriptApp.newTrigger('runOnce').timeBased().after(1000).create();

  // 未処理データの件数を表示
  let count = 0;
  const files = DriveApp.getFolderById(SOURCE_FOLDER_ID).getFiles();
  while (files.hasNext()) {
    const t = files.next().getMimeType();
    if (t.includes('pdf') || t.includes('image')) count++;
  }

  const html =
    '<div style="font-family:-apple-system,sans-serif;padding:24px;text-align:center">' +
    '<p style="font-size:22px;font-weight:bold">処理を開始しました</p>' +
    '<p style="font-size:17px">未処理データ：' + count + '件</p>' +
    '<p style="font-size:15px;color:#666">1〜数分後に「処理済みデータ」をご確認ください。<br>' +
    '件数が多い場合は、残りを5分ごとの自動実行で順次処理します。</p>' +
    '</div>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('スキャン整理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

async function runOnce() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runOnce')
    .forEach(t => ScriptApp.deleteTrigger(t));
  await checkForNewFiles();
}

/**
 * 【補助】判定済みの目印をすべて解除（処理済みデータの全PDFを再判定したい場合のみ実行）
 */
function resetOrientationMarks() {
  let pageToken, count = 0;
  do {
    const res = Drive.Files.list({
      q: `'${DEST_FOLDER_ID}' in parents and trashed = false and mimeType = 'application/pdf'`,
      fields: 'nextPageToken, files(id, appProperties)',
      pageSize: 1000,
      pageToken: pageToken
    });
    (res.files || []).forEach(f => {
      if (!f.appProperties) return;
      Drive.Files.update({ appProperties: { [ORIENT_PROP_KEY]: null, [ORIENT_FAIL_KEY]: null } }, f.id);
      count++;
    });
    pageToken = res.nextPageToken;
  } while (pageToken);
  Logger.log(`目印を解除しました: ${count}件`);
}

/**
 * 【補助】このAPIキーで利用可能なGeminiモデルの一覧をログに出力（代替モデルの確認用）
 */
function listGeminiModels() {
  const apiKey = getApiKey();
  const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1/models?pageSize=1000&key=${apiKey}`, { muteHttpExceptions: true });
  const models = (JSON.parse(res.getContentText()).models || [])
    .filter(m => (m.supportedGenerationMethods || []).includes('generateContent'))
    .map(m => m.name.replace('models/', ''));
  Logger.log(models.join('\n'));
}
