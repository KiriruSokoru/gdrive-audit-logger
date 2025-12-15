/**
 * ==========================================
 * Google Drive Activity Logger (Audit Tool)
 * ==========================================
 * 
 * Автоматически собирает логи активности (Activity API v2) с указанной папки,
 * сопоставляет ID пользователей с реальными именами из справочника
 * и сохраняет события в Google Sheets с разбивкой по дням.
 */

/*************************
 * КОНФИГУРАЦИЯ
 *************************/
// ID корневой папки для аудита (Вставьте свой ID)
const ROOT_FOLDER_ID = 'YOUR_FOLDER_ID_HERE'; 

// Ключ для хранения ID таблицы логов в ScriptProperties
const PROP_KEY_CONFIG_SHEET_ID  = 'driveActivity_configSheetId_v1';
const LOG_SPREADSHEET_NAME_PREFIX = 'DriveActivityLogs_';
const USERS_SHEET_NAME = 'users'; // Имя листа-справочника (ID -> Name)

const MAX_RUNTIME_MS = 5 * 60 * 1000; // Лимит времени работы (5 мин)
const SCRIPT_TZ = Session.getScriptTimeZone();

// Лимиты для защиты от исчерпания квот Google API
const MAX_DRIVE_ACTIVITY_CALLS = 500;
const MAX_DRIVE_FILES_CALLS    = 500;

/*************************
 * ГЛОБАЛЬНЫЕ СЧЁТЧИКИ И КЕШ
 *************************/
let STAT = {
  driveActivityCalls: 0,
  driveFilesCalls: 0,
  errors: 0
};

let USERS_MAP = null;
let FOLDER_CACHE = {}; // Кеш путей: { fileId: { name: "...", parentId: "...", isError: bool } }

/*************************
 * ТОЧКА ВХОДА
 *************************/
function logDriveActivity() {
  const startTime = Date.now();
  console.log('=== Старт logDriveActivity (v3: Smart Cache + Fix) ===');

  // Окно запроса: последние 48 часов (для перекрытия возможных задержек API)
  const now = new Date();
  const fortyEightHoursAgo = new Date(now.getTime() - 48 * 60 * 60 * 1000);
  const fortyEightHoursAgoIso = fortyEightHoursAgo.toISOString();

  // Фильтр: только важные изменения файлов
  const actionFilter = 'detail.action_detail_case:(CREATE EDIT MOVE RENAME)';
  const fullFilter = `time >= "${fortyEightHoursAgoIso}" AND ${actionFilter}`;
  console.log('Полный фильтр: ' + fullFilter);

  // Загружаем мапу пользователей (users)
  const usersMap = getUsersMap_();

  const rowsByDate = {}; // { dateKey: [row,row...] }
  let processedEvents = 0;
  let pageToken = null;

  while (true) {
    // Проверки лимитов времени и квот
    if (Date.now() - startTime > (MAX_RUNTIME_MS - 15000)) {
      console.log('⏳ Почти достигнут лимит времени скрипта, выходим аккуратно.');
      break;
    }
    if (STAT.driveActivityCalls >= MAX_DRIVE_ACTIVITY_CALLS ||
        STAT.driveFilesCalls    >= MAX_DRIVE_FILES_CALLS) {
      console.log('⛔ Достигнут лимит вызовов API (квота), выходим.');
      break;
    }

    const request = {
      pageSize: 100,
      ancestorName: `items/${ROOT_FOLDER_ID}`,
      filter: fullFilter
    };
    if (pageToken) request.pageToken = pageToken;

    let response;
    try {
      STAT.driveActivityCalls++;
      response = DriveActivity.Activity.query(request);
    } catch (e) {
      STAT.errors++;
      console.error('❌ Ошибка DriveActivity.query: ' + e);
      break;
    }

    const activities = response.activities || [];
    console.log(`Получено событий: ${activities.length} (всего обработано: ${processedEvents})`);

    if (!activities.length) {
      pageToken = response.nextPageToken || null;
      if (!pageToken) break;
      continue;
    }

    activities.forEach((activity, idx) => {
      const timeIso = getTimeInfo_(activity) || new Date().toISOString();
      const ts = new Date(timeIso);
      const eventTimeLocal = Utilities.formatDate(ts, SCRIPT_TZ, 'yyyy-MM-dd HH:mm:ss');
      const dateKey = Utilities.formatDate(ts, SCRIPT_TZ, 'yyyy-MM-dd');

      const actors = activity.actors || [];
      const targets = activity.targets || [];
      const primaryDetail = activity.primaryActionDetail || {};
      const actionType = getActionType_(primaryDetail);

      // ---- Получаем инфо об акторе (Жесткая логика) ----
      const actorInfo = getActorInfoForceMap_(actors, usersMap); 
      const targetInfo = getTargetInfo_(targets);

      const fileId   = targetInfo.fileId || '';
      const fileName = targetInfo.title || '';
      const filePath = targetInfo.path || '';

      // Уникальный ключ события для дедупликации
      const eventKey = [
        timeIso,
        fileId,
        actionType,
        actorInfo.personName || ''
      ].join('|');

      // Логируем редко для контроля (сэмпл 1 из 50)
      if ((processedEvents + idx) % 50 === 0) {
        console.log(
          `> #${processedEvents + idx + 1}: ${actionType} | ${actorInfo.name} | ${fileName}`
        );
      }

      if (!rowsByDate[dateKey]) rowsByDate[dateKey] = [];
      rowsByDate[dateKey].push([
        timeIso,                 // 1 eventTimeUTC
        eventTimeLocal,          // 2 eventTimeLocal
        actionType,              // 3 actionType
        actorInfo.name,          // 4 actorName (Читаемое имя)
        actorInfo.personName,    // 5 actorPersonName (ID)
        fileId,                  // 6 fileId
        fileName,                // 7 fileName
        filePath,                // 8 path
        eventKey                 // 9 eventKey
      ]);
    });

    processedEvents += activities.length;

    if (response.nextPageToken) pageToken = response.nextPageToken;
    else break;
  }

  // --- Запись в таблицы
  Object.keys(rowsByDate).forEach(dateKey => {
    const sheet = getOrCreateDailyLogSheet_(dateKey);
    const allRows = rowsByDate[dateKey];
    const existingIndex = loadExistingIndex_(sheet); 

    // 1) Апдейты имен (FORCE) - обновляем имя в старых строках, если оно появилось в справочнике
    const updatesByRow = new Map();

    allRows.forEach(r => {
      const eventKey = r[8];
      const newActorName = r[3]; // Это уже разрешенное имя из users
      
      const hit = existingIndex.get(eventKey);
      if (!hit) return;

      // Если новое имя отличается (например, было ID, стало Имя)
      if (newActorName && newActorName !== hit.actorName) {
        updatesByRow.set(hit.row, newActorName);
      }
    });

    // Оптимизация записи обновлений (батчами)
    const uniqUpdates = Array.from(updatesByRow.entries())
      .map(([row, value]) => ({ row, value }))
      .sort((a, b) => a.row - b.row);

    let applied = 0;
    let i = 0;
    while (i < uniqUpdates.length) {
      let start = i;
      let end = i;
      while (end + 1 < uniqUpdates.length && uniqUpdates[end + 1].row === uniqUpdates[end].row + 1) end++;

      const startRow = uniqUpdates[start].row;
      const block = uniqUpdates.slice(start, end + 1).map(x => [x.value]);
      sheet.getRange(startRow, 4, block.length, 1).setValues(block); // col4 actorName
      applied += block.length;

      i = end + 1;
    }

    // 2) Новые строки (которых нет в existingIndex)
    const newRows = allRows.filter(r => !existingIndex.has(r[8]));
    if (newRows.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    console.log(
      `📅 Дата ${dateKey}: всего событий ${allRows.length}, новых ${newRows.length}, обновлено имён ${applied}`
    );
  });

  const elapsedSec = Math.round((Date.now() - startTime) / 1000);
  console.log('=== Завершено logDriveActivity ===');
  console.log(
    `📊 Статистика:\n` +
    `   Событий обработано: ${processedEvents}\n` +
    `   Drive Activity API: ${STAT.driveActivityCalls}\n` +
    `   Drive Files API:    ${STAT.driveFilesCalls} (кеш сэкономил кучу вызовов)\n` +
    `   Ошибок API:         ${STAT.errors}\n` +
    `   Время выполнения:   ${elapsedSec} сек.`
  );
}
