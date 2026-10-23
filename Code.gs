// Код для Google Apps Script. Вставить целиком в https://script.google.com/
const SPREADSHEET_ID = '1Qo0ee14qCWojHx3bZlaE29A8wGtvmS29zT4WBJi7pSI';
const SHEET_NAME = 'Ответы';
const SEND_EMAIL = true; // письма отправляются владельцу скрипта

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    if (!data.summary || typeof data.summary !== 'string') {
      throw new Error('Нет ответов');
    }
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['Дата', 'Настроение', 'Количество гостей', 'Кого пригласить',
        'Формат праздника', 'Кухня', 'Развлечения', 'Важная деталь',
        'Продолжение вечера', 'Особое пожелание', 'Полное резюме']);
      sheet.setFrozenRows(1);
    }
    const a = data.answers || {};
    const keys = ['Настроение', 'Количество гостей', 'Кого пригласить',
      'Формат праздника', 'Кухня', 'Развлечения', 'Важная деталь',
      'Продолжение вечера', 'Особое пожелание'];
    // Не позволяем тексту ответа превратиться в формулу в Google Sheets.
    const safe = value => {
      const s = String(value == null ? '' : value);
      return /^[=+\-@]/.test(s) ? "'" + s : s;
    };
    sheet.appendRow([new Date(), ...keys.map(k => safe(a[k])), safe(data.summary)]);
    if (SEND_EMAIL) {
      const email = Session.getEffectiveUser().getEmail();
      if (email) {
        try {
          MailApp.sendEmail({to: email, subject: '♡ Ответы: идеальный день рождения',
            body: data.summary});
        } catch (mailError) {
          console.error('Запись сохранена, но письмо не отправлено: ' + mailError);
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
