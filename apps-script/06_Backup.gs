/**
 * 备份。
 *
 * Google Sheet 本身已经有版本记录,但版本记录救不了两种情况:
 *   1. 整个档案被删掉 / 被误清空后过了 30 天
 *   2. 你想看「上个学期结束时」的完整快照
 *
 * 所以这里做的是:定期把整份试算表复制一份到 Google Drive 的备份资料夹,
 * 并且只保留最近 N 份,不会塞爆你的云端硬碟。
 */

const BACKUP_FOLDER_NAME = '补习管理备份';

/** 立刻备份一份 */
function backupNow() {
  const folder = getOrCreateBackupFolder_();
  const ss = ss_();
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HHmm');
  const name = ss.getName() + ' — 备份 ' + stamp;

  DriveApp.getFileById(ss.getId()).makeCopy(name, folder);

  const keep = settingNumber_(getSettings_(), '备份保留份数', 8);
  const removed = pruneBackups_(folder, keep);

  logAudit_('(备份)', '', '建立备份', '', name + (removed ? ',清掉 ' + removed + ' 份旧的' : ''));
  return { name: name, folderUrl: folder.getUrl(), removed: removed };
}

/** 选单用:备份 + 弹窗告诉你放在哪 */
function backupNowWithAlert() {
  const ui = SpreadsheetApp.getUi();
  try {
    const r = backupNow();
    ui.alert('✅ 备份完成',
      '已复制一份:\n' + r.name +
      '\n\n位置:Google 云端硬碟 → ' + BACKUP_FOLDER_NAME +
      (r.removed ? '\n\n(顺便清掉 ' + r.removed + ' 份过旧的备份)' : '') +
      '\n\n资料夹:\n' + r.folderUrl,
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('备份失败', e.message, ui.ButtonSet.OK);
  }
}

/** 装上每周自动备份(每星期日晚上) */
function installWeeklyBackup() {
  const ui = SpreadsheetApp.getUi();
  removeBackupTriggers_();
  ScriptApp.newTrigger('backupNow')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(22)
    .create();
  ui.alert('✅ 已开启自动备份',
    '之后每星期日晚上会自动备份一份到 Google 云端硬碟的「' + BACKUP_FOLDER_NAME + '」资料夹。\n\n' +
    '预设保留最近 8 份,可以在「设置」表改「备份保留份数」。\n\n' +
    '就算你没打开这个试算表,备份还是会跑。',
    ui.ButtonSet.OK);
}

/** 关掉每周自动备份 */
function removeWeeklyBackup() {
  const n = removeBackupTriggers_();
  SpreadsheetApp.getUi().alert(n ? '已关闭自动备份。' : '本来就没有开启自动备份。');
}

function removeBackupTriggers_() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'backupNow') { ScriptApp.deleteTrigger(t); n++; }
  });
  return n;
}

function getOrCreateBackupFolder_() {
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
}

/** 只留最近 keep 份,其余丢进垃圾桶(不是永久删除,还能救回来) */
function pruneBackups_(folder, keep) {
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    files.push({ file: f, created: f.getDateCreated() });
  }
  if (files.length <= keep) return 0;

  files.sort(function (a, b) { return b.created - a.created; }); // 新的在前
  let removed = 0;
  for (let i = keep; i < files.length; i++) {
    files[i].file.setTrashed(true);
    removed++;
  }
  return removed;
}
