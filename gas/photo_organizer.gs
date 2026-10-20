/**
 * Googleドライブ 写真フォルダ整理（撮影日別に自動振り分け）
 *
 * ■ 動作
 *  - 対象フォルダ直下の画像・動画を、撮影日ごとのサブフォルダへ移動する
 *  - 振り分け先フォルダは自動作成（事前作成は不要）
 *  - 日付の判定順：①EXIF撮影日時 → ②ファイル名中の日付（IMG_20240512 等）→ ③ドライブへの作成日
 *  - 6分の実行時間制限に近づいたら自動で続きを予約し、完了したら予約を削除する
 *
 * ■ 事前準備
 *  1. エディタ左の「サービス」＋ → 「Drive API」（v3）を追加
 *  2. 下記 CONFIG.SOURCE_FOLDER_ID を設定
 *  3. まず DRY_RUN = true で organizePhotos を実行し、ログで振り分け先を確認
 *  4. 問題なければ DRY_RUN = false にして再実行
 */

const CONFIG = {
  // 整理対象フォルダのID（URL「drive.google.com/drive/folders/〇〇〇」の〇〇〇部分）
  SOURCE_FOLDER_ID: '●',

  // 振り分け先の作成場所（空欄なら対象フォルダ内に作成）
  DEST_ROOT_FOLDER_ID: '',

  // フォルダ構成
  //  'YEAR_MONTH' → 2024 / 2024-05
  //  'MONTH'      → 2024-05
  //  'YEAR'       → 2024
  //  'DAY'        → 2024 / 2024-05-12
  FOLDER_FORMAT: 'YEAR_MONTH',

  // ファイル名にキーワードを含む場合、日付より優先してこのフォルダへ（不要なら空のまま）
  // 例：{ '桜坂': '桜坂ハイツ', '薬院': '薬院マンション' }
  KEYWORD_RULES: {},

  // true：移動せずログ出力のみ（確認用）
  DRY_RUN: true,

  TIMEZONE: 'Asia/Tokyo',
  TIME_LIMIT_MS: 5 * 60 * 1000,
};

function organizePhotos() {
  const start = Date.now();
  const destRoot = DriveApp.getFolderById(CONFIG.DEST_ROOT_FOLDER_ID || CONFIG.SOURCE_FOLDER_ID);
  const folderCache = {};
  let moved = 0;
  let pageToken = null;

  do {
    const res = Drive.Files.list({
      q: `'${CONFIG.SOURCE_FOLDER_ID}' in parents and trashed = false and ` +
         `(mimeType contains 'image/' or mimeType contains 'video/')`,
      fields: 'nextPageToken, files(id, name, createdTime, imageMediaMetadata/time)',
      pageSize: 200,
      pageToken: pageToken,
    });

    for (const f of res.files || []) {
      if (Date.now() - start > CONFIG.TIME_LIMIT_MS) {
        Logger.log(`時間切れのため中断（${moved}件処理）。1分後に続きを実行します。`);
        if (!CONFIG.DRY_RUN) scheduleContinuation_();
        return;
      }

      const path = resolveDestPath_(f);
      if (CONFIG.DRY_RUN) {
        Logger.log(`[DRY_RUN] ${f.name} → ${path.join('/')}`);
      } else {
        const dest = getOrCreatePath_(destRoot, path, folderCache);
        DriveApp.getFileById(f.id).moveTo(dest);
      }
      moved++;
    }
    pageToken = res.nextPageToken;
  } while (pageToken);

  clearContinuation_();
  Logger.log(`完了：${moved}件${CONFIG.DRY_RUN ? '（DRY_RUN・未移動）' : '移動'}`);
}

/** 振り分け先のフォルダ階層を返す */
function resolveDestPath_(f) {
  for (const keyword in CONFIG.KEYWORD_RULES) {
    if (f.name.indexOf(keyword) !== -1) return [CONFIG.KEYWORD_RULES[keyword]];
  }

  const d = exifDate_(f) || fileNameDate_(f.name) || new Date(f.createdTime);
  const fmt = (p) => Utilities.formatDate(d, CONFIG.TIMEZONE, p);

  switch (CONFIG.FOLDER_FORMAT) {
    case 'MONTH': return [fmt('yyyy-MM')];
    case 'YEAR':  return [fmt('yyyy')];
    case 'DAY':   return [fmt('yyyy'), fmt('yyyy-MM-dd')];
    default:      return [fmt('yyyy'), fmt('yyyy-MM')];
  }
}

/** EXIF撮影日時（"2024:05:12 10:20:33"）を Date に変換 */
function exifDate_(f) {
  const t = f.imageMediaMetadata && f.imageMediaMetadata.time;
  const m = t && t.match(/^(\d{4}):(\d{2}):(\d{2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3], 12) : null;
}

/** ファイル名中の日付（IMG_20240512_xxx / PXL_20240512 / 2024-05-12 など）を Date に変換 */
function fileNameDate_(name) {
  const m = name.match(/(20\d{2})[-_.]?(0[1-9]|1[0-2])[-_.]?(0[1-9]|[12]\d|3[01])/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3], 12) : null;
}

/** フォルダ階層を取得（無ければ作成） */
function getOrCreatePath_(root, names, cache) {
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

function scheduleContinuation_() {
  clearContinuation_();
  ScriptApp.newTrigger('organizePhotos').timeBased().after(60 * 1000).create();
}

function clearContinuation_() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'organizePhotos')
    .forEach((t) => ScriptApp.deleteTrigger(t));
}
