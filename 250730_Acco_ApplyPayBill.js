/**
 * @fileoverview 這是一個 Google Apps Script 範本，用於每月在指定日期自動寄送 Email。
 * 此版本可透過 UI 設定參數、預覽信件、並寄送預覽信，且會自動抓取 Gmail 簽名檔。
 * 修正了在自動觸發器環境下無法使用 UI 的問題。
 * @version 9.0
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
      .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, '設定郵件參數');
}

/**
 * 將使用者從 UI 輸入的設定儲存起來。
 * @param {object} settings - 從前端傳來的設定物件。
 * @returns {string} - 回傳給前端的成功訊息。
 */
function saveSettings(settings) {
  try {
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
        // 如果有設定簽名檔，且 sendAs.displayName 有值，就用它
        if (sendAs && sendAs.displayName) {
            return sendAs.displayName;
        }
        // 否則直接回傳使用者的 Email
        return currentUserEmail;
    } catch (e) {
        // 如果出錯（例如API沒啟用），就回傳 Email 作為備用
        console.error("無法取得預設寄件人名稱: " + e.toString());
        return Session.getActiveUser().getEmail();
    }
}


// =================================================================
// SECTION: 核心寄信邏輯
// =================================================================

/**
 * 【核心修正】寄送每月郵件的函式。
 * 此函式由時間觸發器呼叫，因此移除了所有 UI 互動 (alert)。
 */
function sendMonthlyEmail() {
  const now = new Date();
  const currentMonth = now.getMonth() + 1;
  const currentDay = now.getDate();
  const isNormalMonthSendDay = (currentMonth >= 1 && currentMonth <= 11) && (currentDay === 25);
  const isDecemberSendDay = (currentMonth === 12) && (currentDay === 15);

  // 【修正】如果不是寄信日，就寫入日誌並安靜地結束，而不是跳出 alert。
  if (!isNormalMonthSendDay && !isDecemberSendDay) {
    console.log(`今天 (${currentMonth}/${currentDay}) 不是預定的寄信日，正式信件未寄出。`);
    return; // 直接結束函式執行
  }

  const settings = scriptProperties.getProperties();
  const recipient = settings.recipient;
  
  // 【修正】如果沒有設定收件人，就寫入錯誤日誌，而不是跳出 alert。
  if (!recipient) {
    console.error('錯誤：尚未設定收件者。請透過「郵件自動化工具 > 參數設定」選單進行設定。');
    return; // 設定有誤，結束函式執行
  }

  console.log(`今天是 ${currentMonth}/${currentDay}，為預定寄信日，開始準備正式郵件。`);
  _coreSendEmail(recipient, true); // 傳入 isTriggered = true
}

/**
 * 寄送預覽信給自己。此函式由使用者手動觸發。
 */
function sendPreviewToSelf() {
    const selfEmail = Session.getActiveUser().getEmail();
    if (!selfEmail) {
        SpreadsheetApp.getUi().alert('無法取得您的 Email 地址，無法寄送預覽信。');
        return;
    }
    console.log(`準備寄送預覽信至: ${selfEmail}`);
    _coreSendEmail(selfEmail, false); // 傳入 isTriggered = false
    SpreadsheetApp.getUi().alert(`預覽信件已寄送至您的信箱: ${selfEmail}`);
}

/**
 * 核心的郵件寄送函式。
 * @param {string} recipient - 收件人的 Email 地址。
 * @param {boolean} isTriggered - 判斷此呼叫是否來自自動觸發器，以決定錯誤處理方式。
 */
function _coreSendEmail(recipient, isTriggered) {
  const settings = scriptProperties.getProperties();
  const senderName = settings.senderName;
  const mainContentNormalTemplate = settings.mainContentNormal;
  const mainContentDecemberTemplate = settings.mainContentDecember;

  // 檢查核心設定是否存在
  if (!recipient || !mainContentNormalTemplate || !mainContentDecemberTemplate) {
    const errorMessage = '錯誤：收件者或信件範本尚未設定。請從「參數設定」中填寫。';
    // 【修正】根據呼叫來源決定錯誤提示方式
    if (isTriggered) {
        console.error(errorMessage);
    } else {
        SpreadsheetApp.getUi().alert(errorMessage);
    }
    return;
  }

  const now = new Date();
  const currentMonth = now.getMonth() + 1;
  const currentYear = now.getFullYear();
  const rocYear = currentYear - 1911;
  let deadlineText;
  let mainContent;

  if (currentMonth === 12) {
    const nextRocYear = rocYear + 1;
    deadlineText = `${nextRocYear}年1月5日前截止，遇假日則順延至次一工作日`;
    mainContent = mainContentDecemberTemplate.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{nextRocYear}}/g, nextRocYear);
  } else {
    const nextMonth = currentMonth + 1;
    const deadlineDate = `${rocYear}年${nextMonth}月5日`;
    deadlineText = `${deadlineDate}前截止， 遇假日則順延至次一工作日`;
    mainContent = mainContentNormalTemplate.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{deadlineDate}}/g, deadlineDate);
  }

  const dynamicSubject = `【通知】${rocYear}年${currentMonth}月款項申請(至${deadlineText})`;
  const signature = getGmailSignature();
  const body = `<html><body style="font-family: 'Microsoft JhengHei', sans-serif; font-weight: bold;">${mainContent}${signature}</body></html>`;

  try {
    const mailOptions = { to: recipient, subject: dynamicSubject, htmlBody: body };
    if (senderName) { mailOptions.name = senderName; }
    MailApp.sendEmail(mailOptions);
    console.log("郵件已成功寄送至: " + recipient);
  } catch (e) {
    const errorMessage = "郵件寄送失敗: " + e.toString();
    // 【修正】根據呼叫來源決定錯誤提示方式
    if (isTriggered) {
        console.error(errorMessage);
    } else {
        SpreadsheetApp.getUi().alert(errorMessage);
    }
  }
}

// =================================================================
// SECTION: 輔助函式
// =================================================================

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
    } else {
      console.log('在 Gmail 設定中找不到預設簽名檔。');
      return '';
    }
  } catch (e) {
    console.error(`取得 Gmail 簽名檔時發生錯誤: ${e.message}. 請確認您已在編輯器的「服務」中啟用 Gmail API。`);
    return '';
  }
}

/**
 * 根據範本產生預覽用的 HTML 內容。
 * @param {string} templateHtml - 郵件內容的 HTML 範本。
 * @param {string} templateType - 'normal' 或 'december'。
 * @returns {string} - 替換完變數並加上簽名檔的完整 HTML。
 */
function generatePreviewHtml(templateHtml, templateType) {
  const now = new Date();
  const currentMonth = (templateType === 'december') ? 12 : now.getMonth() + 1;
  const currentYear = now.getFullYear();
  const rocYear = currentYear - 1911;
  let mainContent;

  if (templateType === 'december') {
      const nextRocYear = rocYear + 1;
      mainContent = templateHtml.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{nextRocYear}}/g, nextRocYear);
  } else {
      const nextMonth = currentMonth + 1;
      const deadlineDate = `${rocYear}年${nextMonth}月5日`;
      mainContent = templateHtml.replace(/{{rocYear}}/g, rocYear).replace(/{{currentMonth}}/g, currentMonth).replace(/{{deadlineDate}}/g, deadlineDate);
  }

  const signature = getGmailSignature();
  return mainContent + signature;
}
