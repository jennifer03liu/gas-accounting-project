/**
 * @fileoverview 這是一個 Google Apps Script 範本，用於每月在指定日期自動寄送 Email。
 * 此版本可透過 UI 設定參數、預覽信件、並寄送預覽信，且會自動抓取 Gmail 簽名檔。
 * @version 12.0 (Implemented advanced Markdown-like formatting)
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
  _coreSendEmail(settings.recipient, true);
}

function sendPreviewToSelf() {
    const selfEmail = Session.getActiveUser().getEmail();
    if (!selfEmail) {
        SpreadsheetApp.getUi().alert('無法取得您的 Email 地址，無法寄送預覽信。');
        return;
    }
    console.log(`準備寄送預覽信至: ${selfEmail}`);
    _coreSendEmail(selfEmail, false);
    SpreadsheetApp.getUi().alert(`預覽信件已寄送至您的信箱: ${selfEmail}`);
}

function _coreSendEmail(recipient, isTriggered) {
  try {
    const settings = scriptProperties.getProperties();
    const senderName = settings.senderName;

    const { subject, body } = processEmailTemplates(settings);

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

function processEmailTemplates(settings) {
  const now = new Date();
  const currentMonth = now.getMonth() + 1;
  const currentYear = now.getFullYear();
  const rocYear = currentYear - 1911;

  let subject, body;

  if (currentMonth === 12) {
    subject = settings.subjectDecember || '';
    body = settings.bodyDecember || '';
    const nextRocYear = rocYear + 1;
    subject = subject.replace(/{{nextRocYear}}/g, nextRocYear);
    body = body.replace(/{{nextRocYear}}/g, nextRocYear);
  } else {
    subject = settings.subjectNormal || '';
    body = settings.bodyNormal || '';
    const nextMonth = currentMonth + 1;
    const deadlineDate = `${rocYear}年${nextMonth}月5日`;
    subject = subject.replace(/{{deadlineDate}}/g, deadlineDate);
    body = body.replace(/{{deadlineDate}}/g, deadlineDate);
  }

  subject = subject.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth);
  body = body.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth);

  return { subject, body };
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
  html = html.replace(/\x5B([^\]]+?)\x5D\(([^)]+?)\)/g, '<a href="$2" target="_blank">$1</a>');

  // List items (handles 'l' or '•')
  html = html.replace(/^\s*[l•]\s+(.*)/gm, '<li>$1</li>');
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

function generatePreviewHtml(templateObject, templateType) {
  try {
    const settings = {
      subjectNormal: templateObject.subject, 
      bodyNormal: templateObject.body, 
      subjectDecember: templateObject.subject, 
      bodyDecember: templateObject.body
    };
    
    // Force the month for correct template processing
    const originalDate = Date;
    globalThis.Date = function() {
      if (templateType === 'december') return new originalDate('2023-12-01');
      return new originalDate('2023-01-01');
    };

    const { subject, body } = processEmailTemplates(settings);
    
    globalThis.Date = originalDate; // Restore original Date object

    const finalHtmlBody = markdownToHtml(body);
    const signature = getGmailSignature();
    
    return `<h4>主旨: ${subject}</h4><hr>${finalHtmlBody}${signature}`;
  } catch (e) {
    console.error('產生預覽失敗: ' + e.toString());
    return '產生預覽失敗，請檢查日誌。';
  }
}