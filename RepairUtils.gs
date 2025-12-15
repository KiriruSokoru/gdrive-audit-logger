/*************************
 * СЛУЖЕБНЫЕ (RESET & FORCE FIX)
 *************************/

function resetDriveActivityState() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();

  if (all[PROP_KEY_CONFIG_SHEET_ID]) {
    props.deleteProperty(PROP_KEY_CONFIG_SHEET_ID);
    console.log('Удалён PROP_KEY_CONFIG_SHEET_ID');
  }
  
  // Чистим старые токены если были
  Object.keys(all).forEach(k => {
    if (k.indexOf('driveActivity_pageToken') !== -1) {
      props.deleteProperty(k);
    }
  });
  console.log('Состояние очищено.');
}

/**
 * ЭКСТРЕННЫЙ РЕМОНТ ИМЕН (OFFLINE)
 * Запускать вручную, если нужно обновить старые логи без API,
 * используя только данные из листа users.
 */
function forceFixAllLogNames() {
  console.log('=== ЗАПУСК ПРИНУДИТЕЛЬНОГО РЕМОНТА ИМЕН (OFFLINE) ===');
  const props = PropertiesService.getScriptProperties();
  const configSheetId = props.getProperty(PROP_KEY_CONFIG_SHEET_ID);
  
  if (!configSheetId) {
    console.error('❌ ID таблицы не найден.');
    return;
  }

  const ss = SpreadsheetApp.openById(configSheetId);
  const usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
  
  if (!usersSheet) {
    console.error('❌ Лист users не найден!');
    return;
  }

  const usersData = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 2).getValues();
  const usersMap = new Map();

  usersData.forEach(r => {
    let id = r[0].toString().replace(/\s+/g, '').trim(); 
    let name = r[1].toString().trim();
    if (id && name) {
      usersMap.set(id, name);
      if (id.startsWith('people/')) {
        usersMap.set(id.replace('people/', ''), name);
      }
    }
  });

  const sheets = ss.getSheets();
  let totalFixed = 0;

  sheets.forEach(sheet => {
    if (!sheet.getName().startsWith('log_')) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const range = sheet.getRange(2, 4, lastRow - 1, 2); 
    const values = range.getValues();
    let sheetUpdates = 0;
    const newValues = [];

    for (let i = 0; i < values.length; i++) {
      let currentName = values[i][0];
      let currentIdRaw = values[i][1].toString();
      let cleanIdLog = currentIdRaw.replace(/\s+/g, '').trim();
      let foundName = usersMap.get(cleanIdLog);

      if (foundName && currentName !== foundName) {
        newValues.push([foundName, currentIdRaw]); 
        sheetUpdates++;
      } else {
        newValues.push([currentName, currentIdRaw]);
      }
    }

    if (sheetUpdates > 0) {
      range.setValues(newValues);
      console.log(`⚡ Лист ${sheet.getName()}: исправлено ${sheetUpdates}`);
      totalFixed += sheetUpdates;
    }
  });
  console.log(`\n=== ГОТОВО. Исправлено строк: ${totalFixed} ===`);
}
