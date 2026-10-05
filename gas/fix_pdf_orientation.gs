/**
 * 格納済みPDFの向き補正（縦・横の修正）
 *
 * ・指定フォルダ内のPDFを1件ずつGeminiで向き判定し、必要なページだけ回転させる
 * ・新規作成ではなく「元ファイルを上書き修正」（ファイルID・URL・共有設定はそのまま）
 * ・回転はページの /Rotate 属性の変更のみ（画像の再圧縮・再埋め込みなし）→ 容量はほぼ増えない
 * ・判定済みのファイルには目印（appProperties）を付け、2回目以降は対象外
 *
 * 【事前準備】
 * 1. 「サービス」→「Drive API」を追加（バージョン v3）
 * 2. スクリプトプロパティに GEMINI_API_KEY を設定（既存スクリプトと同じキーで可）
 * 3. ori_setupTrigger() を1回手動実行（10分ごとに自動実行）
 *    ※1回だけ試す場合は ori_fixOrientation() を手動実行
 *
 * ※既存のスキャン整理スクリプトと同じプロジェクトに貼り付けても名前が衝突しないよう、
 *   定数・関数名には ORI_ / ori_ を付けています。
 */

const ORI_TARGET_FOLDER_ID = '1bTo9kY4ZJQCtMeURLJZEHOTIHoyfWJJN'; // マイドライブ/scan/処理済みデータ
const ORI_INCLUDE_SUBFOLDERS = false;   // true にするとサブフォルダ内のPDFも対象

const ORI_PDF_LIB_URL = 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js';
const ORI_PRIMARY_MODEL = 'gemini-3.8-flash';
const ORI_MAX_FALLBACK_MODELS = 2;

// この容量を超えるPDFは処理しない（GASのメモリ・時間制限対策）
const ORI_MAX_BYTES = 5 * 1024 * 1024;  // 5MB
// 失敗がこの回数に達したファイルは以後スキップ
const ORI_MAX_FAILS = 3;

// 判定済みの目印（Driveのファイルに付与する非表示プロパティ）
const ORI_PROP_KEY = 'orientFix';   // 値: done（判定済み）/ large（容量超過）/ error（失敗上限）
const ORI_FAIL_KEY = 'orientFixFails';

async function ori_fixOrientation() {
  const startTime = Date.now();
  const START_CUTOFF = 4 * 60 * 1000;  // 新しいファイルの処理開始は4分以内

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) {
    Logger.log('他の処理が実行中のため、今回はスキップします。');
    return;
  }

  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      Logger.log('エラー: APIキーがスクリプトプロパティに設定されていません。');
      return;
    }

    const targets = ori_listTargets(ORI_TARGET_FOLDER_ID);
    Logger.log(`向き補正の対象（未判定）: ${targets.length}件`);
    let fixed = 0, unchanged = 0;

    for (let i = 0; i < targets.length; i++) {
      if (Date.now() - startTime > START_CUTOFF) {
        Logger.log('実行時間の上限が近いため、残りは次回実行時に処理します。');
        break;
      }
      const f = targets[i];
      const size = Number(f.size || 0);

      if (size > ORI_MAX_BYTES) {
        Logger.log(`【スキップ】容量超過（${(size / 1024 / 1024).toFixed(1)}MB）: ${f.name}`);
        ori_setMark(f.id, { [ORI_PROP_KEY]: 'large' });
        continue;
      }

      Logger.log(`判定開始: ${f.name}`);
      try {
        const result = await ori_processFile(f, apiKey);
        if (result === 'fixed') fixed++;
        else if (result === 'unchanged') unchanged++;
        else ori_countFail(f);
      } catch (err) {
        Logger.log(`【エラー】${f.name}: ${err.toString()}`);
        ori_countFail(f);
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
 * 1ファイル分の処理
 * 戻り値: 'fixed'（回転して上書き）/ 'unchanged'（回転不要）/ null（失敗）
 */
async function ori_processFile(f, apiKey) {
  const blob = DriveApp.getFileById(f.id).getBlob();
  const bytes = blob.getBytes();

  const { PDFDocument, degrees } = ori_loadPdfLib();
  let pdfDoc;
  try {
    pdfDoc = await PDFDocument.load(new Uint8Array(bytes), { ignoreEncryption: true, updateMetadata: false });
  } catch (err) {
    Logger.log(`PDFの読み込みに失敗しました: ${err.toString()}`);
    return null;
  }
  const pageCount = pdfDoc.getPageCount();

  const rotations = ori_requestRotations(apiKey, Utilities.base64Encode(bytes), pageCount);
  if (!rotations) return null;
  Logger.log(`回転判定: [${rotations.join(', ')}]`);

  // 回転不要 → ファイルには触れず、判定済みの目印のみ付与
  if (!rotations.some(r => r !== 0)) {
    ori_setMark(f.id, { [ORI_PROP_KEY]: 'done', [ORI_FAIL_KEY]: null });
    Logger.log('補正不要');
    return 'unchanged';
  }

  // 必要なページの /Rotate 属性のみ変更（画像データはそのまま）
  pdfDoc.getPages().forEach((page, i) => {
    const r = rotations[i] || 0;
    if (r === 0) return;
    const current = page.getRotation().angle || 0;
    page.setRotation(degrees((current + r) % 360));
  });
  const out = await pdfDoc.save({ useObjectStreams: true });
  const newBlob = Utilities.newBlob(Array.from(out), MimeType.PDF, f.name);

  // 元ファイルを上書き（ファイルID・URL・共有設定・ファイル名は維持）
  Drive.Files.update(
    { appProperties: { [ORI_PROP_KEY]: 'done', [ORI_FAIL_KEY]: null } },
    f.id,
    newBlob
  );
  Logger.log(`【補正完了】${f.name}（${(bytes.length / 1024).toFixed(0)}KB → ${(out.length / 1024).toFixed(0)}KB）`);
  return 'fixed';
}

/**
 * Geminiで各ページの回転角度を取得（モデル切替＋リトライ）
 */
function ori_requestRotations(apiKey, base64, pageCount) {
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

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'application/pdf', data: base64 } }] }],
      generationConfig: { responseMimeType: 'application/json' }
    }),
    muteHttpExceptions: true
  };

  const models = ori_getModelCandidates(apiKey);
  const maxRetries = 3;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const model = models[(attempt - 1) % models.length];
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    try {
      const res = UrlFetchApp.fetch(url, options);
      const code = res.getResponseCode();
      const body = res.getContentText();

      if (code >= 200 && code < 300) {
        const json = JSON.parse(body);
        const parts = (json.candidates && json.candidates[0].content && json.candidates[0].content.parts) || [];
        const text = parts.map(p => p.text || '').join('').replace(/```(json)?/g, '');
        const m = text.match(/\{[\s\S]*\}/);
        if (!m) { Logger.log(`APIレスポンスの形式が想定外です: ${body}`); return null; }
        const raw = JSON.parse(m[0]).rotations;
        if (!Array.isArray(raw)) { Logger.log(`APIレスポンスの形式が想定外です: ${body}`); return null; }
        const rotations = [];
        for (let i = 0; i < pageCount; i++) {
          const r = Number(raw[i]) || 0;
          rotations.push((((Math.round(r / 90) * 90) % 360) + 360) % 360);
        }
        Logger.log(`使用モデル: ${model}`);
        return rotations;
      }

      Logger.log(`【エラー】API呼び出し失敗 (ステータス: ${code} / モデル: ${model}) - 試行: ${attempt}/${maxRetries}`);
      if (attempt === maxRetries) return null;
      const nextModel = models[attempt % models.length];
      Utilities.sleep(nextModel !== model ? 5000 : attempt * 20000);
    } catch (err) {
      Logger.log(`通信エラーまたはJSONパースエラー: ${err.toString()}`);
      return null;
    }
  }
  return null;
}

/**
 * 未判定のPDFを一覧化（判定済み・容量超過・失敗上限のものは除外）
 */
function ori_listTargets(folderId) {
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
          if (ORI_INCLUDE_SUBFOLDERS) folders.push(f.id);
          return;
        }
        if ((f.appProperties || {})[ORI_PROP_KEY]) return;
        results.push(f);
      });
      pageToken = res.nextPageToken;
    } while (pageToken);
  }
  results.sort((a, b) => new Date(a.createdTime) - new Date(b.createdTime));  // 古い順
  return results;
}

/**
 * 目印（appProperties）の付与 ※ファイル本体は変更しない
 */
function ori_setMark(fileId, props) {
  Drive.Files.update({ appProperties: props }, fileId);
}

function ori_countFail(f) {
  const fails = Number((f.appProperties || {})[ORI_FAIL_KEY] || 0) + 1;
  if (fails >= ORI_MAX_FAILS) {
    Logger.log(`失敗が${fails}回に達したため、以後この判定はスキップします: ${f.name}`);
    ori_setMark(f.id, { [ORI_PROP_KEY]: 'error' });
  } else {
    ori_setMark(f.id, { [ORI_FAIL_KEY]: String(fails) });
  }
}

/**
 * 使用するモデルの候補リスト（メイン＋flash系の代替モデル、6時間キャッシュ）
 */
function ori_getModelCandidates(apiKey) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'ori_models_' + ORI_PRIMARY_MODEL;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  let list = [ORI_PRIMARY_MODEL];
  try {
    const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models?pageSize=1000&key=${apiKey}`, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const version = n => { const m = n.match(/gemini-(\d+(?:\.\d+)?)/); return m ? parseFloat(m[1]) : 0; };
      const fallbacks = (JSON.parse(res.getContentText()).models || [])
        .filter(m => (m.supportedGenerationMethods || []).includes('generateContent'))
        .map(m => m.name.replace('models/', ''))
        .filter(n => /flash/.test(n) && !/(image|tts|audio|live|embedding|exp)/.test(n) && n !== ORI_PRIMARY_MODEL)
        .sort((a, b) =>
          (/preview/.test(a) - /preview/.test(b)) ||
          (version(b) - version(a)) ||
          (/lite/.test(a) - /lite/.test(b)));
      list = list.concat([...new Set(fallbacks)].slice(0, ORI_MAX_FALLBACK_MODELS));
    }
  } catch (e) {
    Logger.log(`モデル一覧の取得に失敗したため、メインモデルのみで実行します: ${e.toString()}`);
  }
  cache.put(cacheKey, JSON.stringify(list), 6 * 60 * 60);
  return list;
}

/**
 * pdf-libの読み込み
 */
function ori_loadPdfLib() {
  if (typeof PDFLib !== 'undefined') return PDFLib;
  const setTimeout = function (f, t) { Utilities.sleep(t || 0); return f(); };
  eval(UrlFetchApp.fetch(ORI_PDF_LIB_URL).getContentText());
  return PDFLib;
}

/**
 * 定期実行トリガーの設定（最初に1回だけ手動実行）
 */
function ori_setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'ori_fixOrientation')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('ori_fixOrientation').timeBased().everyMinutes(10).create();
}

/**
 * 【補助】判定済みの目印をすべて解除（全ファイルを再判定したい場合のみ実行）
 */
function ori_resetMarks() {
  let pageToken, count = 0;
  do {
    const res = Drive.Files.list({
      q: `'${ORI_TARGET_FOLDER_ID}' in parents and trashed = false and mimeType = 'application/pdf'`,
      fields: 'nextPageToken, files(id, appProperties)',
      pageSize: 1000,
      pageToken: pageToken
    });
    (res.files || []).forEach(f => {
      if (!f.appProperties) return;
      Drive.Files.update({ appProperties: { [ORI_PROP_KEY]: null, [ORI_FAIL_KEY]: null } }, f.id);
      count++;
    });
    pageToken = res.nextPageToken;
  } while (pageToken);
  Logger.log(`目印を解除しました: ${count}件`);
}
