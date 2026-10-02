/**
 * 写真台帳 自動作成（Google Drive → Excel / PDF）
 *
 * ルートフォルダ直下に「案件フォルダ」を作り、写真を入れるだけで
 * 案件フォルダ内の「納品」フォルダに 写真台帳.xlsx / 写真台帳.pdf を出力する。
 *
 *  案件フォルダ名：  物件名_工事名        例）桜坂ハイツ101_原状回復工事
 *  写真はファイル名順に No.1〜 を採番（並び順を指定したい場合はファイル名先頭に 01_ 等を付ける）
 */

const CONFIG = {
  ROOT_FOLDER_ID: '●',          // 写真台帳ルートフォルダのID（URLの folders/ 以降）
  OUTPUT_SUBFOLDER: '納品',
  STABLE_MINUTES: 10,            // 最後の写真追加からこの分数経過後に作成（アップロード途中の作成防止）
  PHOTOS_PER_PAGE: 3,
  COMPANY: '株式会社デイトン',
  NOTIFY_EMAIL: '',              // 作成完了を通知するアドレス（空欄なら通知なし）
  THUMB_SIZE: 1200,              // 貼付け画像の長辺px（容量対策）
};

const LAYOUT = {
  COL_WIDTHS: [60, 650],         // A:No. B:写真
  HEADER_ROWS: 4,
  BLOCK_H: 280,
  GAP_H: 10,
};

/** 初回に1回だけ実行：10分おきの自動チェックを登録 */
function setup() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'scan')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scan').timeBased().everyMinutes(10).create();
}

/** 手動で全案件を作り直す */
function rebuildAll() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  scan(true);
}

function scan(force) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    const folders = root.getFolders();
    const started = Date.now();
    while (folders.hasNext()) {
      if (Date.now() - started > 4 * 60 * 1000) break; // 実行時間上限(6分)対策。残りは次回
      const folder = folders.next();
      const photos = listPhotos_(folder);
      if (!photos.length) continue;

      const sig = signature_(photos);
      const key = 'sig_' + folder.getId();
      if (props.getProperty(key) === sig) continue;

      const latest = Math.max(...photos.map(f => f.getLastUpdated().getTime()));
      if (force !== true && Date.now() - latest < CONFIG.STABLE_MINUTES * 60000) continue;

      buildLedger_(folder, photos);
      props.setProperty(key, sig);
    }
  } finally {
    lock.releaseLock();
  }
}

function listPhotos_(folder) {
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getMimeType().indexOf('image/') === 0) files.push(f);
  }
  return files.sort((a, b) =>
    a.getName().localeCompare(b.getName(), 'ja', { numeric: true }));
}

function signature_(photos) {
  const src = photos
    .map(f => [f.getId(), f.getName(), f.getLastUpdated().getTime()].join(':'))
    .join('|');
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, src, Utilities.Charset.UTF_8));
}

function buildLedger_(folder, photos) {
  const folderName = folder.getName();
  const idx = folderName.indexOf('_');
  const property = idx > 0 ? folderName.slice(0, idx) : folderName;
  const work = idx > 0 ? folderName.slice(idx + 1) : '';
  const baseName = folderName + '_写真台帳';

  const ss = SpreadsheetApp.create(baseName);
  const ssFile = DriveApp.getFileById(ss.getId());
  try {
    const per = CONFIG.PHOTOS_PER_PAGE;
    const pageCount = Math.ceil(photos.length / per);
    for (let p = 0; p < pageCount; p++) {
      const sh = p === 0 ? ss.getSheets()[0] : ss.insertSheet();
      sh.setName('P' + (p + 1));
      layoutPage_(sh, property, work, p + 1, pageCount);
      photos.slice(p * per, (p + 1) * per)
        .forEach((file, i) => fillBlock_(sh, i, p * per + i + 1, file));
    }
    SpreadsheetApp.flush();

    const out = getOrCreateSubfolder_(folder, CONFIG.OUTPUT_SUBFOLDER);
    const xlsx = out.createFile(exportBlob_(ss.getId(), 'xlsx', baseName + '.xlsx'));
    const pdf = out.createFile(exportBlob_(ss.getId(), 'pdf', baseName + '.pdf'));
    trashOld_(out, baseName, [xlsx.getId(), pdf.getId()]);

    if (CONFIG.NOTIFY_EMAIL) {
      MailApp.sendEmail(CONFIG.NOTIFY_EMAIL, '【写真台帳】' + folderName + ' 作成完了',
        '写真 ' + photos.length + ' 枚で写真台帳を作成しました。\n\n' +
        'Excel：' + xlsx.getUrl() + '\nPDF：' + pdf.getUrl() + '\n');
    }
  } finally {
    ssFile.setTrashed(true);
  }
}

function layoutPage_(sh, property, work, page, pageCount) {
  const L = LAYOUT;
  const totalRows = L.HEADER_ROWS + CONFIG.PHOTOS_PER_PAGE * 2;
  sh.getRange(1, 1, totalRows, 2).setFontFamily('Noto Sans JP').setFontSize(10).setVerticalAlignment('middle');
  if (sh.getMaxColumns() > 2) sh.deleteColumns(3, sh.getMaxColumns() - 2);
  if (sh.getMaxRows() > totalRows) sh.deleteRows(totalRows + 1, sh.getMaxRows() - totalRows);
  L.COL_WIDTHS.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setHiddenGridlines(true);

  sh.setRowHeight(1, 40);
  sh.getRange('A1:B1').merge().setValue('写 真 台 帳').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setRowHeight(2, 24);
  sh.getRange('A2:B2').merge().setValue('物件名：' + property + (work ? '　／　工事名：' + work : ''));
  sh.setRowHeight(3, 24);
  sh.getRange('A3:B3').merge()
    .setValue(CONFIG.COMPANY + '　作成日：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd') +
      '　' + page + ' / ' + pageCount)
    .setHorizontalAlignment('right');
  sh.setRowHeight(4, 8);

  for (let i = 0; i < CONFIG.PHOTOS_PER_PAGE; i++) {
    const r = blockTop_(i);
    sh.setRowHeight(r, L.BLOCK_H);
    sh.setRowHeight(r + 1, L.GAP_H);
    sh.getRange(r, 1).setHorizontalAlignment('center').setFontSize(12).setFontWeight('bold');
    sh.getRange(r, 1, 1, 2)
      .setBorder(true, true, true, true, true, false, '#666666', SpreadsheetApp.BorderStyle.SOLID);
  }
}

function blockTop_(i) {
  return LAYOUT.HEADER_ROWS + 1 + i * 2;
}

function fillBlock_(sh, i, no, file) {
  const L = LAYOUT;
  const r = blockTop_(i);
  sh.getRange(r, 1).setValue(no);

  const boxW = L.COL_WIDTHS[1] - 12;
  const boxH = L.BLOCK_H - 12;
  try {
    const img = sh.insertImage(imageBlob_(file), 2, r);
    const scale = Math.min(boxW / img.getWidth(), boxH / img.getHeight());
    const w = Math.round(img.getWidth() * scale);
    const h = Math.round(img.getHeight() * scale);
    img.setWidth(w).setHeight(h)
      .setAnchorCellXOffset(Math.round((L.COL_WIDTHS[1] - w) / 2))
      .setAnchorCellYOffset(Math.round((L.BLOCK_H - h) / 2));
  } catch (e) {
    sh.getRange(r, 2).setValue('画像取得失敗：' + file.getName()).setHorizontalAlignment('center');
    console.warn(file.getName(), e);
  }
}

/** Drive API v3 でサムネイルURLを取得 */
function thumbnailLink_(id) {
  const url = 'https://www.googleapis.com/drive/v3/files/' + id +
    '?fields=thumbnailLink&supportsAllDrives=true';
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  return res.getResponseCode() === 200 ? JSON.parse(res.getContentText()).thumbnailLink : null;
}

/** 縮小画像を取得（スマホ写真は insertImage の上限2MBを超えるため）。HEICもJPEG化される */
function imageBlob_(file) {
  const link = thumbnailLink_(file.getId());
  if (link) {
    const res = UrlFetchApp.fetch(link.replace(/=s\d+$/, '=s' + CONFIG.THUMB_SIZE), {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() === 200) return res.getBlob();
  }
  if (file.getSize() < 2 * 1024 * 1024) return file.getBlob();
  throw new Error('サムネイル未生成かつ2MB超');
}

function exportBlob_(id, format, name) {
  let url = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=' + format;
  if (format === 'pdf') {
    url += '&size=A4&portrait=true&scale=4&gridlines=false&printtitle=false' +
      '&sheetnames=false&fzr=false&horizontal_alignment=CENTER' +
      '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4';
  }
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error(format + ' 出力失敗: HTTP ' + res.getResponseCode());
  }
  return res.getBlob().setName(name);
}

function getOrCreateSubfolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** 再作成時、同名の旧ファイルをゴミ箱へ */
function trashOld_(folder, baseName, keepIds) {
  ['.xlsx', '.pdf'].forEach(ext => {
    const it = folder.getFilesByName(baseName + ext);
    while (it.hasNext()) {
      const f = it.next();
      if (keepIds.indexOf(f.getId()) < 0) f.setTrashed(true);
    }
  });
}
