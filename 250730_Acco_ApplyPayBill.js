/**
 * @fileoverview 這是一個 Google Apps Script 範本，用於每月在指定日期自動寄送 Email。
 * 此版本可透過 UI 設定參數、預覽信件、並寄送預覽信，且會自動抓取 Gmail 簽名檔。
 * @version 16.0
 * @changelog
 * - 發信日提前判斷 25 號是否為假日，若是假日則提前到最近的工作日
 * - 假日資料快取到 Properties，定期（如每日）更新一次
 * - 統一錯誤提示方式，全部記錄到 log，不彈窗
 * - 增加 processEmailTemplates、calculateDeadline 等邏輯的單元測試（見 Tests.js）
 */

// =================================================================
// SECTION: UI 與設定管理
// =================================================================

const scriptProperties = PropertiesService.getScriptProperties();

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('郵件自動化工具')
      .addItem('參數設定', 'showSettingsDialog')
      .addSeparator()
      .addItem('執行正式寄信', 'sendMonthlyEmail')
      .addItem('預覽寄送給自己', 'sendPreviewToSelf')
      .addToUi();
}

function showSettingsDialog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('SettingsUI')
        .setWidth(850) 
        .setHeight(750); // Increased height for new format instructions
    SpreadsheetApp.getUi().showModalDialog(html, '設定郵件參數');
  } catch (e) {
    console.error('showSettingsDialog 失敗: ' + e.toString());
    SpreadsheetApp.getUi().alert('無法開啟設定介面，請檢查日誌。');
  }
}

function saveSettings(settings) {
  try {
    scriptProperties.setProperties(settings);
    console.log('設定已儲存:', settings);
    return '設定已成功儲存！';
  } catch (e) {
    logError(e, 'saveSettings');
    return '儲存失敗，請檢查日誌。';
  }
}

function getSettings() {
  try {
    return {
      properties: scriptProperties.getProperties(),
      defaultSenderName: getDefaultSenderName()
    };
  } catch (e) {
    logError(e, 'getSettings');
    return { properties: {}, defaultSenderName: '' };
  }
}

function getDefaultSenderName() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    const sendAs = Gmail.Users.Settings.SendAs.get('me', currentUserEmail);
    return (sendAs && sendAs.displayName) ? sendAs.displayName : currentUserEmail;
  } catch (e) {
    logError(e, 'getDefaultSenderName');
    return Session.getActiveUser().getEmail() || '';
  }
}

// =================================================================
// SECTION: 假日快取管理
// =================================================================

function updateHolidayCache() {
  try {
    const url = "https://calendar.google.com/calendar/ical/zh-tw.taiwan%23holiday%40group.v.calendar.google.com/public/basic.ics";
    const icalData = UrlFetchApp.fetch(url).getContentText();
    const parsed = parseHolidayData(icalData);
    scriptProperties.setProperty('holidayCache', JSON.stringify(parsed));
    scriptProperties.setProperty('holidayCacheUpdated', new Date().toISOString());
    console.log('假日資料已更新快取');
  } catch (e) {
    logError(e, 'updateHolidayCache');
  }
}

function getCachedHolidayData() {
  const cache = scriptProperties.getProperty('holidayCache');
  if (cache) {
    try {
      return JSON.parse(cache);
    } catch (e) {
      logError(e, 'getCachedHolidayData');
      return { holidays: [], workdays: [] };
    }
  }
  // 若無快取則即時抓取並快取
  try {
    updateHolidayCache();
    const cache2 = scriptProperties.getProperty('holidayCache');
    if (cache2) return JSON.parse(cache2);
  } catch (e) {
    logError(e, 'getCachedHolidayData-fetch');
  }
  return { holidays: [], workdays: [] };
}

// =================================================================
// SECTION: 發信日計算
// =================================================================

function getSendDate(year, month, holidays, workdays) {
  holidays = Array.isArray(holidays) ? holidays : [];
  workdays = Array.isArray(workdays) ? workdays : [];
  let date = new Date(year, month - 1, 25);
  while (true) {
    const time = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
    let dayOfWeek = date.getDay();
    // 若是假日且不是補班日，往前一天
    if ((dayOfWeek === 0 || dayOfWeek === 6 || holidays.indexOf(time) !== -1) && workdays.indexOf(time) === -1) {
      date.setDate(date.getDate() - 1);
      // 如果已經跨月，代表整月都沒有可用的工作日
      if (date.getMonth() + 1 !== month) {
        return null;
      }
    } else {
      break;
    }
  }
  // 檢查是否有效日期
  if (isNaN(date.getTime())) {
    return null;
  }
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

// =================================================================
// SECTION: 核心寄信邏輯
// =================================================================

function sendMonthlyEmail() {
  try {
    const now = new Date();
    const year = now.getFullYear();
    const month = now.getMonth() + 1;
    const holidayData = getCachedHolidayData();
    const sendDate = getSendDate(year, month, holidayData.holidays, holidayData.workdays);

    if (!sendDate) {
      logError('本月無有效寄信日（可能整月都是假日或資料異常）。', 'sendMonthlyEmail');
      return;
    }

    // 只在發信日當天寄信
    if (
      now.getFullYear() === sendDate.getFullYear() &&
      now.getMonth() === sendDate.getMonth() &&
      now.getDate() === sendDate.getDate()
    ) {
      const settings = scriptProperties.getProperties();
      if (!settings.recipient) {
        logError('尚未設定收件者。請透過「郵件自動化工具 > 參數設定」選單進行設定。', 'sendMonthlyEmail');
        return;
      }
      console.log(`今天是 ${year}/${month}/${now.getDate()}，為本月發信日，開始準備正式郵件。`);
      _coreSendEmail(settings.recipient, true, year, month);
    } else {
      console.log(`今天 (${year}/${month}/${now.getDate()}) 不是本月發信日 (${sendDate.getFullYear()}/${sendDate.getMonth()+1}/${sendDate.getDate()})，正式信件未寄出。`);
    }
  } catch (e) {
    logError(e, 'sendMonthlyEmail');
  }
}

function sendPreviewToSelf() {
  try {
    const selfEmail = Session.getActiveUser().getEmail();
    if (!selfEmail) {
      logError('無法取得您的 Email 地址，無法寄送預覽信。', 'sendPreviewToSelf');
      return;
    }
    // 彈窗選擇年份和月份
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('請輸入年份(例如2024)及月份(1-12)', '格式: YYYY-MM，例如: 2024-12', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;
    const match = response.getResponseText().match(/^(\d{4})-(\d{1,2})$/);
    if (!match) {
      logError('格式錯誤，請輸入正確的年份及月份 (例如: 2024-12)', 'sendPreviewToSelf');
      return;
    }
    const year = parseInt(match[1], 10);
    const month = parseInt(match[2], 10);
    if (month < 1 || month > 12) {
      logError('月份必須介於 1 到 12', 'sendPreviewToSelf');
      return;
    }
    console.log(`準備寄送預覽信至: ${selfEmail}, 年份: ${year}, 月份: ${month}`);
    _coreSendEmail(selfEmail, false, year, month);
    // 不再用 alert
    console.log(`預覽信件已寄送至您的信箱: ${selfEmail}`);
  } catch (e) {
    logError(e, 'sendPreviewToSelf');
  }
}

function _coreSendEmail(recipient, isTriggered, year, month) {
  try {
    const settings = scriptProperties.getProperties();
    const senderName = settings.senderName;
    const { subject, body } = processEmailTemplates(settings, year, month);

    if (!recipient || !subject || !body) {
      throw new Error('收件者、信件主旨或內文範本尚未設定。');
    }

    const finalHtmlBody = markdownToHtml(body);
    const signature = getGmailSignature();
    const fullBody = `<html><body>${finalHtmlBody}${signature}</body></html>`;

    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: fullBody,
      name: senderName
    });

    console.log("郵件已成功寄送至: " + recipient);

  } catch (e) {
    logError(e, '_coreSendEmail');
  }
}

// =================================================================
// SECTION: 輔助函式
// =================================================================

function processEmailTemplates(settings, year, month) {
  let now;
  if (year && month) {
    now = new Date(year, month - 1, 1);
  } else {
    now = new Date();
  }
  const currentMonth = now.getMonth() + 1;
  const currentYear = now.getFullYear();
  const rocYear = currentYear - 1911;

  let subject, body;
  const holidayData = getCachedHolidayData();

  if (currentMonth === 12) {
    subject = settings.subjectDecember || '';
    body = settings.bodyDecember || '';
    const nextRocYear = rocYear + 1;
    const deadlineDate = calculateDeadline(currentYear, 12, holidayData.holidays, holidayData.workdays);
    subject = subject.replace(/{{nextRocYear}}/g, nextRocYear);
    body = body.replace(/{{nextRocYear}}/g, nextRocYear);
    subject = subject.replace(/{{deadlineDate}}/g, deadlineDate);
    body = body.replace(/{{deadlineDate}}/g, deadlineDate);
  } else {
    subject = settings.subjectNormal || '';
    body = settings.bodyNormal || '';
    const deadlineDate = calculateDeadline(currentYear, currentMonth, holidayData.holidays, holidayData.workdays);
    subject = subject.replace(/{{deadlineDate}}/g, deadlineDate);
    body = body.replace(/{{deadlineDate}}/g, deadlineDate);
  }

  subject = subject.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth);
  body = body.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth);

  return { subject, body };
}

function getAndParseHolidayData() {
  // 已廢棄，改用 getCachedHolidayData
  return getCachedHolidayData();
}

function parseHolidayData(icalData) {
  const holidays = [];
  const workdays = [];
  const events = icalData.split('BEGIN:VEVENT');
  
  events.forEach(event => {
    if (!event.includes('DTSTART')) return;

    const summaryMatch = event.match(/SUMMARY:(.+)/);
    const descriptionMatch = event.match(/DESCRIPTION:(.+)/);
    const dtstartMatch = event.match(/DTSTART;VALUE=DATE:(\d{8})/);

    if (dtstartMatch && summaryMatch && descriptionMatch) {
      const dateStr = dtstartMatch[1];
      const summary = summaryMatch[1].trim();
      const description = descriptionMatch[1].trim();
      
      const date = new Date(parseInt(dateStr.substring(0, 4)), parseInt(dateStr.substring(4, 6)) - 1, parseInt(dateStr.substring(6, 8)));

      if (summary.includes('補班')) {
        workdays.push(date.getTime());
      } else if (description.includes('國定假日') || summary.includes('補假') || summary.includes('厂礼拜')) {
        holidays.push(date.getTime());
      }
    }
  });

  return { holidays, workdays };
}

function calculateDeadline(year, month, holidays, workdays) {
  let deadline = new Date(year, month, 5);
  while (true) {
    const deadlineTime = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate()).getTime();
    let dayOfWeek = deadline.getDay();
    if (workdays.indexOf(deadlineTime) !== -1) {
      break;
    }
    if ((dayOfWeek === 6 || dayOfWeek === 0) || holidays.indexOf(deadlineTime) !== -1) {
      deadline.setDate(deadline.getDate() + 1);
    } else {
      break;
    }
  }
  const rocYear = deadline.getFullYear() - 1911;
  const nextMonth = deadline.getMonth() + 1;
  const day = deadline.getDate();
  return `${rocYear}年${nextMonth}月${day}日`;
}

function markdownToHtml(text) {
  if (!text) return '';

  let html = text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  // Custom Tags - Order is important
  // Specific colored tags first
  html = html.replace(/\*\*紅字\*\*(.+?)\*\*紅字\*\*/g, '<span style="background-color:#ffff00; color:#cc0000; font-weight:bold;">$1</span>');
  html = html.replace(/\*\*黃底\*\*(.+?)\*\*黃底\*\*/g, '<span style="background-color:#ffff00; color:#000000; font-weight:bold;">$1</span>');
  
  // General bold tag
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  
  // Links
  html = html.replace(/\x5B([^\\]+?)\x5D\(([^)]+?)\)/g, '<a href="$2" target="_blank">$1</a>');

  // List items (handles 'l', '•', or '-')
  html = html.replace(/^\s*[l•-]\s+(.*)/gm, '<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>\s*)+/g, '<ul>$&</ul>');

  // Newlines
  html = html.replace(/\r\n|\n|\r/g, '<br>\n');
  html = html.replace(/<br>\n<ul>/g, '<ul>').replace(/<\/ul><br>\n/g, '</ul>');

  return html;
}

function getGmailSignature() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    const sendAs = Gmail.Users.Settings.SendAs.get('me', currentUserEmail);
    return (sendAs && sendAs.signature) ? sendAs.signature : '';
  } catch (e) {
    console.error(`取得 Gmail 簽名檔時發生錯誤: ${e.message}`);
    return '';
  }
}

function generatePreviewHtml(templateObject, templateType, year, month) {
  try {
    const settings = {
      subjectNormal: templateObject.subject, 
      bodyNormal: templateObject.body, 
      subjectDecember: templateObject.subject, 
      bodyDecember: templateObject.body
    };
    // 預覽時讓使用者選擇年份和月份
    let previewYear, previewMonth;
    if (year && month) {
      previewYear = year;
      previewMonth = month;
    } else {
      const now = new Date();
      previewYear = now.getFullYear();
      if (templateType === 'december') {
        previewMonth = 12;
      } else {
        previewMonth = now.getMonth() + 1;
        if (previewMonth === 12) previewMonth = 11;
      }
    }
    const { subject, body } = processEmailTemplates(settings, previewYear, previewMonth);

    const finalHtmlBody = markdownToHtml(body);
    const signature = getGmailSignature();
    return `<h4>主旨: ${subject}</h4><hr>${finalHtmlBody}${signature}`;
  } catch (e) {
    logError(e, 'generatePreviewHtml');
    return '產生預覽失敗，請檢查日誌。';
  }
}

function logError(e, context) {
  const msg = `[${context}] ${e && e.message ? e.message : e}`;
  console.error(msg);
  if (e && e.stack) console.error(e.stack);
  Logger.log(msg);
}

function testShowSendDate(year, month) {
  const holidayData = getCachedHolidayData();
  const sendDate = getSendDate(year, month, holidayData.holidays, holidayData.workdays);
  if (!sendDate) {
    console.log(`本月(${year}/${month})無有效寄信日`);
  } else {
    console.log(`本月(${year}/${month})寄信日：${sendDate.getFullYear()}/${sendDate.getMonth()+1}/${sendDate.getDate()}`);
  }
}

//範例：查詢 2025 年 12 月寄信日
//testShowSendDate(2025, 12);
