/**
 * @fileoverview 這是一個 Google Apps Script 範本，用於每月在指定日期自動寄送 Email。
 * 此版本可透過 UI 設定參數、預覽信件、並寄送預覽信，且會自動抓取 Gmail 簽名檔。
 * 修正了在自動觸發器環境下無法使用 UI 的問題。
 * @version 10.0 (Refactored to separate content from presentation)
 */

// =================================================================
// SECTION: UI 與設定管理
// =================================================================

const scriptProperties = PropertiesService.getScriptProperties();

/**
 * 當試算表檔案被開啟時，自動執行的觸發器，用來建立自訂選單。
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('郵件自動化工具')
      .addItem('參數設定', 'showSettingsDialog')
      .addSeparator()
      .addItem('執行正式寄信', 'sendMonthlyEmail')
      .addItem('預覽寄送給自己', 'sendPreviewToSelf')
      .addToUi();
}

/**
 * 顯示參數設定的對話框。
 */
function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SettingsUI')
      .setWidth(850) 
      .setHeight(600); // Increased height for new fields
  SpreadsheetApp.getUi().showModalDialog(html, '設定郵件參數');
}

/**
 * 將使用者從 UI 輸入的設定儲存起來。
 * @param {object} settings - 從前端傳來的設定物件。
 * @returns {string} - 回傳給前端的成功訊息。
 */
function saveSettings(settings) {
  try {
    // Remove old properties to avoid confusion
    scriptProperties.deleteProperty('mainContentNormal');
    scriptProperties.deleteProperty('mainContentDecember');
    
    scriptProperties.setProperties(settings);
    console.log('設定已儲存:', settings);
    return '設定已成功儲存！';
  } catch (e) {
    console.error('儲存設定失敗:', e);
    return '儲存失敗，請檢查日誌。';
  }
}

/**
 * 讀取已儲存的設定，並額外取得預設寄件人名稱。
 * @returns {object} - 回傳包含設定屬性與預設名稱的物件。
 */
function getSettings() {
  const settings = {
    properties: scriptProperties.getProperties(),
    defaultSenderName: getDefaultSenderName()
  };
  return settings;
}

/**
 * 取得 Gmail 預設的寄件人名稱。
 * @returns {string} - 返回預設的寄件人名稱，若無則返回 Email 地址。
 */
function getDefaultSenderName() {
    try {
        const currentUserEmail = Session.getActiveUser().getEmail();
        const sendAs = Gmail.Users.Settings.SendAs.get('me', currentUserEmail);
        if (sendAs && sendAs.displayName) {
            return sendAs.displayName;
        }
        return currentUserEmail;
    } catch (e) {
        console.error("無法取得預設寄件人名稱: " + e.toString());
        return Session.getActiveUser().getEmail();
    }
}


// =================================================================
// SECTION: 核心寄信邏輯
// =================================================================

/**
 * 寄送每月郵件的函式。
 */
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
  const recipient = settings.recipient;
  
  if (!recipient) {
    console.error('錯誤：尚未設定收件者。請透過「郵件自動化工具 > 參數設定」選單進行設定。');
    return;
  }

  console.log(`今天是 ${currentMonth}/${currentDay}，為預定寄信日，開始準備正式郵件。`);
  _coreSendEmail(recipient, true);
}

/**
 * 寄送預覽信給自己。
 */
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

/**
 * 【核心修改】核心的郵件寄送函式。
 * @param {string} recipient - 收件人的 Email 地址。
 * @param {boolean} isTriggered - 判斷此呼叫是否來自自動觸發器。
 */
function _coreSendEmail(recipient, isTriggered) {
  const settings = scriptProperties.getProperties();
  const senderName = settings.senderName;

  const now = new Date();
  const currentMonth = now.getMonth() + 1;
  const currentYear = now.getFullYear();
  const rocYear = currentYear - 1911;
  
  let subjectTemplate, bodyTemplate, deadlineText;

  if (currentMonth === 12) {
    subjectTemplate = settings.subjectDecember;
    bodyTemplate = settings.bodyDecember;
    const nextRocYear = rocYear + 1;
    deadlineText = `${nextRocYear}年1月5日前截止，遇假日則順延至次一工作日`;
    // Replace placeholders
    subjectTemplate = subjectTemplate.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{nextRocYear}}/g, nextRocYear);
    bodyTemplate = bodyTemplate.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{nextRocYear}}/g, nextRocYear);
  } else {
    subjectTemplate = settings.subjectNormal;
    bodyTemplate = settings.bodyNormal;
    const nextMonth = currentMonth + 1;
    const deadlineDate = `${rocYear}年${nextMonth}月5日`;
    deadlineText = `${deadlineDate}前截止， 遇假日則順延至次一工作日`;
    // Replace placeholders
    subjectTemplate = subjectTemplate.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{deadlineDate}}/g, deadlineDate);
    bodyTemplate = bodyTemplate.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{deadlineDate}}/g, deadlineDate);
  }

  if (!recipient || !subjectTemplate || !bodyTemplate) {
    const errorMessage = '錯誤：收件者、信件主旨或內文範本尚未設定。請從「參數設定」中填寫。';
    if (isTriggered) { console.error(errorMessage); } 
    else { SpreadsheetApp.getUi().alert(errorMessage); }
    return;
  }

  // Convert body from Markdown to HTML
  const finalHtmlBody = markdownToHtml(bodyTemplate);
  const signature = getGmailSignature();
  const fullBody = `<html><body>${finalHtmlBody}${signature}</body></html>`;

  try {
    const mailOptions = { to: recipient, subject: subjectTemplate, htmlBody: fullBody };
    if (senderName) { mailOptions.name = senderName; }
    MailApp.sendEmail(mailOptions);
    console.log("郵件已成功寄送至: " + recipient);
  } catch (e) {
    const errorMessage = "郵件寄送失敗: " + e.toString();
    if (isTriggered) { console.error(errorMessage); } 
    else { SpreadsheetApp.getUi().alert(errorMessage); }
  }
}

// =================================================================
// SECTION: 輔助函式
// =================================================================

/**
 * Converts a simple Markdown-like text to HTML.
 * @param {string} text The plain text to convert.
 * @returns {string} The converted HTML string.
 */
function markdownToHtml(text) {
  if (!text) return '';

  let html = text
    // Escape HTML characters to prevent XSS, but keep brackets for links
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  // Handle Markdown links [text](url) before other processing
  html = html.replace(/\ \[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank">$1</a>');

  // Handle list items (lines starting with - or *)
  html = html.replace(/^\s*[-*]\s+(.*)/gm, '<li>$1</li>');
  html = html.replace(/(<li>.*<\/li>\s*)+/g, '<ul>$&</ul>');

  // Handle newlines - convert them to <br>
  html = html.replace(/\n/g, '<br>\n');

  // Clean up <br> tags around lists
  html = html.replace(/<br>\n<ul>/g, '<ul>');
  html = html.replace(/<\/ul><br>\n/g, '</ul>');

  return html;
}

/**
 * 取得 Gmail 帳號設定的預設簽名檔。
 * @returns {string} - HTML 格式的簽名檔字串。
 */
function getGmailSignature() {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    const sendAs = Gmail.Users.Settings.SendAs.get('me', currentUserEmail);
    if (sendAs && sendAs.signature) {
      console.log('成功取得 Gmail 簽名檔。');
      return sendAs.signature;
    }
  } catch (e) {
    console.error(`取得 Gmail 簽名檔時發生錯誤: ${e.message}. 請確認您已在編輯器的「服務」中啟用 Gmail API。`);
    return '';
  }
}

/**
 * 【核心修改】根據範本產生預覽用的 HTML 內容。
 * @param {object} templateObject - 包含 subject 和 body 的物件。
 * @param {string} templateType - 'normal' 或 'december'。
 * @returns {string} - 替換完變數並加上簽名檔的完整 HTML。
 */
function generatePreviewHtml(templateObject, templateType) {
  let { subject, body } = templateObject;

  const now = new Date();
  const currentMonth = (templateType === 'december') ? 12 : now.getMonth() + 1;
  const currentYear = now.getFullYear();
  const rocYear = currentYear - 1911;

  if (templateType === 'december') {
      const nextRocYear = rocYear + 1;
      subject = subject.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{nextRocYear}}/g, nextRocYear);
      body = body.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{nextRocYear}}/g, nextRocYear);
  } else {
      const nextMonth = currentMonth + 1;
      const deadlineDate = `${rocYear}年${nextMonth}月5日`;
      subject = subject.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{deadlineDate}}/g, deadlineDate);
      body = body.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{deadlineDate}}/g, deadlineDate);
  }

  // Convert body from Markdown to HTML for preview
  const finalHtmlBody = markdownToHtml(body);
  const signature = getGmailSignature();
  
  // For preview, let's include the subject in the body so the user can see it.
  return `<h4>主旨: ${subject}</h4><hr>${finalHtmlBody}${signature}`;
}