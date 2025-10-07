/**
 * @fileoverview 這是一個 Google Apps Script 範本，用於每月在指定日期自動寄送 Email。
 * 此版本可透過 UI 設定參數、預覽信件、並寄送預覽信，且會自動抓取 Gmail 簽名檔。
 * @version 15.0 (Added year/month selection for preview, unified email content generation)
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
    console.error('儲存設定失敗: ' + e.toString());
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
    console.error('取得設定失敗: ' + e.toString());
    return { properties: {}, defaultSenderName: '' }; // Return empty object on failure
  }
}

function getDefaultSenderName() {
    try {
        const currentUserEmail = Session.getActiveUser().getEmail();
        const sendAs = Gmail.Users.Settings.SendAs.get('me', currentUserEmail);
        return (sendAs && sendAs.displayName) ? sendAs.displayName : currentUserEmail;
    } catch (e) {
        console.error("無法取得預設寄件人名稱: " + e.toString());
        return Session.getActiveUser().getEmail() || '';
    }
}

// =================================================================
// SECTION: 核心寄信邏輯
// =================================================================

function sendMonthlyEmail() {
  const now = new Date();
  const currentMonth = now.getMonth() + 1;
  const currentDay = now.getDate();
  const isNormalMonthSendDay = (currentMonth >= 1 && currentMonth <= 11) && (currentDay === 25);
  const isDecemberSendDay = (currentMonth === 12) && (currentDay === 15);

  if (!isNormalMonthSendDay && !isDecemberSendDay) {
    console.log(`今天 (${currentMonth}/${currentDay}) 不是預定的寄信日，正式信件未寄出。`);
    return;
  }

  const settings = scriptProperties.getProperties();
  if (!settings.recipient) {
    console.error('錯誤：尚未設定收件者。請透過「郵件自動化工具 > 參數設定」選單進行設定。');
    return;
  }

  console.log(`今天是 ${currentMonth}/${currentDay}，為預定寄信日，開始準備正式郵件。`);
  const nowDate = new Date();
  const year = nowDate.getFullYear();
  const month = nowDate.getMonth() + 1;
  _coreSendEmail(settings.recipient, true, year, month);
}

function sendPreviewToSelf() {
    const selfEmail = Session.getActiveUser().getEmail();
    if (!selfEmail) {
        SpreadsheetApp.getUi().alert('無法取得您的 Email 地址，無法寄送預覽信。');
        return;
    }
    // 彈窗選擇年份和月份
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('請輸入年份(例如2024)及月份(1-12)', '格式: YYYY-MM，例如: 2024-12', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;
    const match = response.getResponseText().match(/^(\d{4})-(\d{1,2})$/);
    if (!match) {
      ui.alert('格式錯誤，請輸入正確的年份及月份 (例如: 2024-12)');
      return;
    }
    const year = parseInt(match[1], 10);
    const month = parseInt(match[2], 10);
    if (month < 1 || month > 12) {
      ui.alert('月份必須介於 1 到 12');
      return;
    }
    console.log(`準備寄送預覽信至: ${selfEmail}, 年份: ${year}, 月份: ${month}`);
    _coreSendEmail(selfEmail, false, year, month);
    ui.alert(`預覽信件已寄送至您的信箱: ${selfEmail}`);
}

function _coreSendEmail(recipient, isTriggered, year, month) {
  try {
    const settings = scriptProperties.getProperties();
    const senderName = settings.senderName;

    // 使用指定的年份和月份產生信件內容
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
    console.error("郵件寄送失敗: " + e.toString());
    if (!isTriggered) {
      SpreadsheetApp.getUi().alert("郵件寄送失敗: " + e.message);
    }
  }
}

// =================================================================
// SECTION: 輔助函式
// =================================================================

function processEmailTemplates(settings, year, month) {
  // year/month 可選，預設用現在
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
  const holidayData = getAndParseHolidayData();

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
  const url = "https://calendar.google.com/calendar/ical/zh-tw.taiwan%23holiday%40group.v.calendar.google.com/public/basic.ics";
  try {
    const icalData = UrlFetchApp.fetch(url).getContentText();
    return parseHolidayData(icalData);
  } catch (e) {
    console.error("無法獲取或解析日曆資料: " + e.toString());
    // Return empty arrays on failure to avoid breaking the script
    return { holidays: [], workdays: [] };
  }
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
  // Note: month is 1-based
  let deadline = new Date(year, month, 5); // month is 0-based in Date constructor, so month becomes next month

  while (true) {
    const deadlineTime = new Date(deadline.getFullYear(), deadline.getMonth(), deadline.getDate()).getTime();
    let dayOfWeek = deadline.getDay();

    // If it's a workday, we can stop, even if it's a weekend.
    if (workdays.indexOf(deadlineTime) !== -1) {
      break;
    }

    // If it's a weekend (and not a designated workday) or a holiday, we increment the date.
    if ((dayOfWeek === 6 || dayOfWeek === 0) || holidays.indexOf(deadlineTime) !== -1) {
      deadline.setDate(deadline.getDate() + 1);
    } else {
      break; // It's a working day
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
    console.error('產生預覽失敗: ' + e.toString());
    return '產生預覽失敗，請檢查日誌。';
  }
}
