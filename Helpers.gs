/*************************
 * ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
 *************************/

function getTimeInfo_(activity) {
  if (activity.timestamp) return activity.timestamp;
  if (activity.timeRange) {
    if (activity.timeRange.endTime) return activity.timeRange.endTime;
    if (activity.timeRange.startTime) return activity.timeRange.startTime;
  }
  return null;
}

function getActionType_(detail) {
  if (!detail) return '';
  for (var key in detail) {
    if (!detail.hasOwnProperty(key)) continue;
    return key;
  }
  return '';
}

/**
 * НОРМАЛИЗАЦИЯ (Агрессивная)
 * Удаляет все пробелы (в т.ч. неразрывные) для корректного сравнения ключей
 */
function normalizePeopleKey_(s) {
  if (!s) return '';
  return s.toString().replace(/\s+/g, '').trim();
}

function getKnownUserFromActors_(actors) {
  if (!actors || !actors.length) return { personName: '', isCurrentUser: false };
  for (const a of actors) {
    const known = a && a.user && a.user.knownUser ? a.user.knownUser : null;
    if (known && known.personName) {
      return { 
        personName: normalizePeopleKey_(known.personName),
        isCurrentUser: !!known.isCurrentUser 
      };
    }
  }
  return { personName: '', isCurrentUser: false };
}

function getActorInfoForceMap_(actors, usersMap) {
  const ku = getKnownUserFromActors_(actors);
  const personName = ku.personName || ''; 

  let name = '';
  
  if (personName && usersMap && usersMap.hasOwnProperty(personName)) {
    name = usersMap[personName];
  } 
  else if (ku.isCurrentUser) {
    name = 'me';
  } 
  else {
    name = personName;
  }

  return { name, personName };
}

function getTargetInfo_(targets) {
  if (!targets || !targets.length) return {};
  const t = targets[0];
  const driveItem = t.driveItem || {};
  let fileId = '';
  let title = '';

  if (driveItem) {
    title = driveItem.title || '';
    if (driveItem.name && driveItem.name.indexOf('items/') === 0) {
      fileId = driveItem.name.replace('items/', '');
    }
  }

  let path = '';
  if (fileId) {
    try {
      path = buildPathFromRoot_(fileId);
    } catch (e) {
      STAT.errors++;
      path = '(Ошибка получения пути)';
    }
  }

  return { fileId, title, path };
}

/**
 * УМНОЕ ПОСТРОЕНИЕ ПУТИ С КЕШЕМ
 * Экономит вызовы API, запоминая родителей папок
 */
function buildPathFromRoot_(fileId) {
  const pathParts = [];
  let currentId = fileId;
  let depth = 0; 
  
  while (currentId && depth < 20) {
    depth++;
    
    // 1. Проверяем кеш
    if (FOLDER_CACHE[currentId]) {
      const cached = FOLDER_CACHE[currentId];
      if (cached.isError) return cached.name + (pathParts.length ? '/' + pathParts.join('/') : '');
      
      pathParts.unshift(cached.name);
      if (!cached.parentId || cached.parentId === ROOT_FOLDER_ID) break;
      currentId = cached.parentId;
      continue;
    }

    // 2. Запрос к API
    try {
      STAT.driveFilesCalls++;
      const file = Drive.Files.get(currentId, { fields: 'id,name,parents,trashed' });
      
      const name = file.name;
      const parents = file.parents || [];
      const parentId = parents.length > 0 ? parents[0] : null;

      FOLDER_CACHE[currentId] = { name: name, parentId: parentId, isError: false };
      
      pathParts.unshift(name);

      if (parentId === ROOT_FOLDER_ID || !parentId) {
        break;
      }
      currentId = parentId;

    } catch (e) {
      const errorMsg = e.toString();
      let statusLabel = '(Нет доступа)';
      if (errorMsg.indexOf('File not found') !== -1 || errorMsg.indexOf('404') !== -1) {
        statusLabel = '(Удалён/Не найден)';
      }
      
      FOLDER_CACHE[currentId] = { name: statusLabel, parentId: null, isError: true };
      pathParts.unshift(statusLabel);
      break;
    }
  }

  return pathParts.join('/');
}

/*************************
 * ЛОГ-ТАБЛИЦЫ
 *************************/
function getOrCreateDailyLogSheet_(dateKey) {
  const props = PropertiesService.getScriptProperties();
  let configSheetId = props.getProperty(PROP_KEY_CONFIG_SHEET_ID);

  let spreadsheet;
  if (configSheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(configSheetId);
    } catch (e) {
      STAT.errors++;
      console.log('⚠️ Не удалось открыть лог, создаём новый: ' + e);
      spreadsheet = createLogSpreadsheet_();
      props.setProperty(PROP_KEY_CONFIG_SHEET_ID, spreadsheet.getId());
    }
  } else {
    console.log('📝 Лог-файл ещё не создан, создаём новый.');
    spreadsheet = createLogSpreadsheet_();
    props.setProperty(PROP_KEY_CONFIG_SHEET_ID, spreadsheet.getId());
  }

  const sheetName = `log_${dateKey}`;
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    // console.log('Создаём новый лист: ' + sheetName);
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow([
      'eventTimeUTC',
      'eventTimeLocal',
      'actionType',
      'actorName',
      'actorPersonName',
      'fileId',
      'fileName',
      'path',
      'eventKey'
    ]);
  }
  return sheet;
}

function createLogSpreadsheet_() {
  const year = new Date().getFullYear();
  const name = `${LOG_SPREADSHEET_NAME_PREFIX}${year}`;
  console.log('Creating new Spreadsheet: ' + name);
  const ss = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(ss.getId());
  
  // Перемещение файла в целевую папку (опционально)
  try {
     const targetFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
     file.moveTo(targetFolder);
  } catch (e) {
     console.log('⚠️ Не удалось переместить лог-файл в целевую папку. Он останется в корне Drive.');
  }
  return ss;
}

function loadExistingIndex_(sheet) {
  const lastRow = sheet.getLastRow();
  const index = new Map();
  if (lastRow < 2) return index;

  const values = sheet.getRange(2, 4, lastRow - 1, 6).getValues();
  for (let i = 0; i < values.length; i++) {
    const actorName = values[i][0] || '';
    const actorPersonName = values[i][1] || '';
    const eventKey = values[i][5] || '';
    if (eventKey) index.set(eventKey, { row: i + 2, actorName, actorPersonName });
  }
  return index;
}

/*************************
 * СПРАВОЧНИК ПОЛЬЗОВАТЕЛЕЙ
 *************************/
function getUsersMap_() {
  if (USERS_MAP !== null) return USERS_MAP;

  const props = PropertiesService.getScriptProperties();
  const configSheetId = props.getProperty(PROP_KEY_CONFIG_SHEET_ID);
  if (!configSheetId) {
    console.log('⚠️ Справочник недоступен (нет ID файла).');
    USERS_MAP = {};
    return USERS_MAP;
  }

  let ss;
  try {
    ss = SpreadsheetApp.openById(configSheetId);
  } catch (e) {
    STAT.errors++;
    console.log('❌ Ошибка открытия таблицы users: ' + e);
    USERS_MAP = {};
    return USERS_MAP;
  }

  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    console.log('⚠️ Лист "' + USERS_SHEET_NAME + '" не найден, создаём пустой.');
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.appendRow(['actorPersonName', 'displayName']);
    USERS_MAP = {};
    return USERS_MAP;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    USERS_MAP = {};
    return USERS_MAP;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};
  
  values.forEach(r => {
    // Агрессивная нормализация
    const rawKey = (r[0] || '');
    const personName = normalizePeopleKey_(rawKey);
    const displayName = (r[1] || '').toString().trim();
    
    if (personName && displayName) {
      map[personName] = displayName;
    }
  });

  USERS_MAP = map;
  console.log(`✅ Загружено users: ${Object.keys(USERS_MAP).length} записей`);
  return USERS_MAP;
}
