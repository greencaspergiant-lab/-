/**
 * Googleドライブ 写真整理（Geminiで件名付け＋分類フォルダへ振り分け）
 *
 * ■ 動作
 *  - 対象フォルダ直下の画像をGeminiに見せ、「分類」と「件名」を判定させる
 *  - ファイル名を「撮影日_件名.拡張子」に変更し、分類フォルダへ移動（フォルダは自動作成）
 *    例：IMG_1234.jpg → 浴室 / 2026-10 / 20261020_浴室換気扇カバー破損.jpg
 *  - 判定に失敗した画像は対象フォルダに残す（再実行で再判定）
 *  - 6分の実行時間制限に近づいたら自動で続きを予約し、完了したら予約を削除する
 *  - photo_organizer.gs と同じプロジェクトに入れても名前は衝突しない
 *
 * ■ 事前準備
 *  1. Google AI Studio（https://aistudio.google.com/apikey）でAPIキーを発行
 *  2. GASの「プロジェクトの設定」→「スクリプト プロパティ」に
 *     プロパティ名：GEMINI_API_KEY ／ 値：発行したキー を追加
 *  3. エディタ左の「サービス」＋ → 「Drive API」（v3）を追加
 *  4. 下記 AI_CONFIG.SOURCE_FOLDER_ID を設定
 *  5. まず DRY_RUN = true で organizePhotosAI を実行し、ログで判定結果を確認
 *  6. 問題なければ DRY_RUN = false にして再実行
 */

const AI_CONFIG = {
  // 整理対象フォルダのID（URL「drive.google.com/drive/folders/〇〇〇」の〇〇〇部分）
  SOURCE_FOLDER_ID: '●',

  // 振り分け先の作成場所（空欄なら対象フォルダ内に作成）
  DEST_ROOT_FOLDER_ID: '',

  // 分類（＝フォルダ名）。Geminiはこの中から必ず1つを選ぶ。自由に追加・変更可
  CATEGORIES: [
    '外観',
    '共用部',
    '居室',
    'キッチン',
    '浴室',
    'トイレ',
    '洗面所',
    '設備・不具合',
    '工事',
    '書類・図面',
    'その他',
  ],

  // 物件・用途など、判定のヒント（空欄可）
  CONTEXT: '賃貸マンションの管理・原状回復・修繕で撮影した写真',

  // フォルダ構成
  //  'CATEGORY'       → 浴室
  //  'CATEGORY_MONTH' → 浴室 / 2026-10
  //  'MONTH_CATEGORY' → 2026-10 / 浴室
  FOLDER_FORMAT: 'CATEGORY_MONTH',

  // 件名の最大文字数
  TITLE_MAX_LENGTH: 20,

  // 使用モデル（廃止・変更時はここだけ差し替え）
  MODEL: 'gemini-flash-latest',

  // 1枚ごとの待機時間（無料枠のレート制限対策。有料枠なら 0 で可）
  WAIT_MS: 4000,

  // true：判定結果をログ出力のみ（名前変更・移動はしない）
  DRY_RUN: true,

  TIMEZONE: 'Asia/Tokyo',
  TIME_LIMIT_MS: 5 * 60 * 1000,
};

function organizePhotosAI() {
  const start = Date.now();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('スクリプト プロパティ GEMINI_API_KEY が未設定です');

  const destRoot = DriveApp.getFolderById(AI_CONFIG.DEST_ROOT_FOLDER_ID || AI_CONFIG.SOURCE_FOLDER_ID);
  const folderCache = {};
  let done = 0;
  let failed = 0;
  let pageToken = null;

  do {
    const res = Drive.Files.list({
      q: `'${AI_CONFIG.SOURCE_FOLDER_ID}' in parents and trashed = false and mimeType contains 'image/'`,
      fields: 'nextPageToken, files(id, name, mimeType, createdTime, thumbnailLink, imageMediaMetadata/time)',
      pageSize: 100,
      pageToken: pageToken,
    });

    for (const f of res.files || []) {
      if (Date.now() - start > AI_CONFIG.TIME_LIMIT_MS) {
        Logger.log(`時間切れのため中断（成功${done}件／失敗${failed}件）。1分後に続きを実行します。`);
        // 1件も成功していない場合（APIキー誤り・上限到達等）は続きを予約しない
        if (!AI_CONFIG.DRY_RUN && done > 0) aiScheduleContinuation_();
        return;
      }

      try {
        const result = aiClassify_(f, apiKey);
        const date = aiPhotoDate_(f);
        const newName = `${Utilities.formatDate(date, AI_CONFIG.TIMEZONE, 'yyyyMMdd')}_${result.title}${aiExt_(f.name)}`;
        const path = aiDestPath_(result.category, date);

        if (AI_CONFIG.DRY_RUN) {
          Logger.log(`[DRY_RUN] ${f.name} → ${path.join('/')}/${newName}`);
        } else {
          const file = DriveApp.getFileById(f.id);
          file.setName(newName);
          file.moveTo(aiGetOrCreatePath_(destRoot, path, folderCache));
        }
        done++;
      } catch (e) {
        Logger.log(`[失敗] ${f.name}：${e.message}`);
        failed++;
      }
      if (AI_CONFIG.WAIT_MS) Utilities.sleep(AI_CONFIG.WAIT_MS);
    }
    pageToken = res.nextPageToken;
  } while (pageToken);

  aiClearContinuation_();
  Logger.log(`完了：成功${done}件／失敗${failed}件${AI_CONFIG.DRY_RUN ? '（DRY_RUN・未変更）' : ''}`);
}

/** Geminiで分類と件名を判定 */
function aiClassify_(f, apiKey) {
  const image = aiImageForGemini_(f);
  const prompt = [
    `次の写真を分類し、件名を付けてください。用途：${AI_CONFIG.CONTEXT}`,
    `・category：候補から最も当てはまるものを1つ。判断できなければ「その他」`,
    `・title：写っている場所・物・状態がわかる日本語の件名（${AI_CONFIG.TITLE_MAX_LENGTH}文字以内、空白や記号なし）`,
    `  例：浴室換気扇カバー破損／洋室クロス剥がれ／エントランス外観／給湯器型番ラベル`,
  ].join('\n');

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: image.mimeType, data: Utilities.base64Encode(image.bytes) } },
      ],
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          category: { type: 'STRING', enum: AI_CONFIG.CATEGORIES },
          title: { type: 'STRING' },
        },
        required: ['category', 'title'],
      },
    },
  };

  const res = UrlFetchApp.fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${AI_CONFIG.MODEL}:generateContent`,
    {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    }
  );
  if (res.getResponseCode() !== 200) {
    throw new Error(`Gemini API ${res.getResponseCode()}：${res.getContentText().slice(0, 200)}`);
  }

  const text = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
  const json = JSON.parse(text);
  const category = AI_CONFIG.CATEGORIES.indexOf(json.category) !== -1 ? json.category : 'その他';
  const title = aiSanitize_(json.title).slice(0, AI_CONFIG.TITLE_MAX_LENGTH) || '件名なし';
  return { category: category, title: title };
}

/** 判定用の画像を取得（容量・通信量削減のため縮小サムネイルを優先） */
function aiImageForGemini_(f) {
  if (f.thumbnailLink) {
    const url = f.thumbnailLink.replace(/=s\d+$/, '=s1024');
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() === 200) {
      const blob = res.getBlob();
      return { mimeType: blob.getContentType() || 'image/jpeg', bytes: blob.getBytes() };
    }
  }
  const blob = DriveApp.getFileById(f.id).getBlob();
  return { mimeType: f.mimeType, bytes: blob.getBytes() };
}

/** 撮影日：EXIF → ファイル名中の日付 → ドライブ作成日 */
function aiPhotoDate_(f) {
  const t = f.imageMediaMetadata && f.imageMediaMetadata.time;
  let m = t && t.match(/^(\d{4}):(\d{2}):(\d{2})/);
  if (!m) m = f.name.match(/(20\d{2})[-_.]?(0[1-9]|1[0-2])[-_.]?(0[1-9]|[12]\d|3[01])/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3], 12) : new Date(f.createdTime);
}

function aiDestPath_(category, date) {
  const month = Utilities.formatDate(date, AI_CONFIG.TIMEZONE, 'yyyy-MM');
  switch (AI_CONFIG.FOLDER_FORMAT) {
    case 'CATEGORY':       return [category];
    case 'MONTH_CATEGORY': return [month, category];
    default:               return [category, month];
  }
}

/** ファイル名に使えない文字・空白を除去 */
function aiSanitize_(s) {
  return String(s || '').replace(/[\\/:*?"<>|\s　]/g, '');
}

function aiExt_(name) {
  const m = name.match(/\.[^.]+$/);
  return m ? m[0] : '';
}

/** フォルダ階層を取得（無ければ作成） */
function aiGetOrCreatePath_(root, names, cache) {
  let folder = root;
  let key = '';
  for (const name of names) {
    key += '/' + name;
    if (!cache[key]) {
      const it = folder.getFoldersByName(name);
      cache[key] = it.hasNext() ? it.next() : folder.createFolder(name);
    }
    folder = cache[key];
  }
  return folder;
}

function aiScheduleContinuation_() {
  aiClearContinuation_();
  ScriptApp.newTrigger('organizePhotosAI').timeBased().after(60 * 1000).create();
}

function aiClearContinuation_() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'organizePhotosAI')
    .forEach((t) => ScriptApp.deleteTrigger(t));
}
