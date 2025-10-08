/**
 * @fileoverview Unit tests for the Gas-Accounting-Project.
 * To run these tests, open the Apps Script editor, select the 'runAllTests' function, and click 'Run'.
 * Then, check the logs (View > Logs) for the results.
 */

// =================================================================
// SECTION: Test Runner & Assertions
// =================================================================

/**
 * A simple testing function to compare two values.
 * @param {string} testName The name of the test.
 * @param {*} expected The expected value.
 * @param {*} actual The actual value.
 */
function assertEquals(testName, expected, actual) {
  if (expected === actual) {
    console.log(`✅ PASSED: ${testName}`);
  } else {
    console.error(`❌ FAILED: ${testName}`);
    console.error(`  --> Expected: ${expected}`);
    console.error(`  --> Actual:   ${actual}`);
  }
}

function assertDeepEquals(testName, expected, actual) {
  const expStr = JSON.stringify(expected);
  const actStr = JSON.stringify(actual);
  if (expStr === actStr) {
    console.log(`✅ PASSED: ${testName}`);
  } else {
    console.error(`❌ FAILED: ${testName}`);
    console.error(`  --> Expected: ${expStr}`);
    console.error(`  --> Actual:   ${actStr}`);
  }
}

/**
 * Main function to run all unit tests.
 */
function runAllTests() {
  console.log('=============== Running All Unit Tests ===============');
  
  // Run tests for the markdownToHtml function
  test_markdownToHtml_convertsNewlines();
  test_markdownToHtml_convertsLinks();
  test_markdownToHtml_convertsListItems();
  test_markdownToHtml_handlesMixedContent();
  test_markdownToHtml_handlesEmptyInput();

  // 新增 processEmailTemplates/calculateDeadline 測試
  test_calculateDeadline_basic();
  test_calculateDeadline_holiday();
  test_calculateDeadline_workday();
  test_getSendDate_basic();
  test_getSendDate_holiday();
  test_processEmailTemplates_variables();

  console.log('==================== Test Run Complete ====================');
}


// =================================================================
// SECTION: Test Cases
// =================================================================

function test_markdownToHtml_convertsNewlines() {
  const input = 'Hello\nWorld';
  const expected = 'Hello<br>\nWorld<br>\n';
  const actual = markdownToHtml(input);
  assertEquals('markdownToHtml: Should convert newlines to <br> tags', expected, actual);
}

function test_markdownToHtml_convertsLinks() {
  const input = 'Check out [Google](https://google.com)';
  const expected = 'Check out <a href="https://google.com" target="_blank">Google</a><br>\n';
  const actual = markdownToHtml(input);
  assertEquals('markdownToHtml: Should convert Markdown links to HTML anchor tags', expected, actual);
}

function test_markdownToHtml_convertsListItems() {
  const input = '- First item\n- Second item';
  const expected = '<ul><li>First item</li>\n<li>Second item</li>\n</ul>';
  const actual = markdownToHtml(input);
  assertEquals('markdownToHtml: Should convert dashed list items to an HTML <ul> list', expected, actual);
}

function test_markdownToHtml_handlesMixedContent() {
  const input = 'Here is a list:\n- One\n- Two\nAnd a [link](https://example.com).';
  const expected = 'Here is a list:<br>\n<ul><li>One</li>\n<li>Two</li>\n</ul>And a <a href="https://example.com" target="_blank">link</a>.<br>\n';
  const actual = markdownToHtml(input);
  assertEquals('markdownToHtml: Should handle a mix of lists, links, and newlines', expected, actual);
}

function test_markdownToHtml_handlesEmptyInput() {
  const input = '';
  const expected = '';
  const actual = markdownToHtml(input);
  assertEquals('markdownToHtml: Should return an empty string for empty input', expected, actual);
}

function test_calculateDeadline_basic() {
  // 2024/8/5 (週一) 非假日
  const holidays = [];
  const workdays = [];
  const result = calculateDeadline(2024, 7, holidays, workdays); // 8月
  assertEquals('calculateDeadline: 基本週一', '113年8月5日', result);
}

function test_calculateDeadline_holiday() {
  // 2024/8/5 (週一) 是假日，應往後推一天
  const holidays = [new Date(2024, 7, 5).getTime()];
  const workdays = [];
  const result = calculateDeadline(2024, 7, holidays, workdays);
  assertEquals('calculateDeadline: 假日往後推', '113年8月6日', result);
}

function test_calculateDeadline_workday() {
  // 2024/8/5 (週一) 是假日，但也是補班日，應不推
  const holidays = [new Date(2024, 7, 5).getTime()];
  const workdays = [new Date(2024, 7, 5).getTime()];
  const result = calculateDeadline(2024, 7, holidays, workdays);
  assertEquals('calculateDeadline: 補班日不推', '113年8月5日', result);
}

function test_getSendDate_basic() {
  // 2024/8/25 (週日) 非假日，應往前推到 23 號 (週五)
  const holidays = [];
  const workdays = [];
  const result = getSendDate(2024, 8, holidays, workdays);
  assertEquals('getSendDate: 週日往前推', new Date(2024, 7, 23).toISOString().slice(0,10), result.toISOString().slice(0,10));
}

function test_getSendDate_holiday() {
  // 2024/8/25 (週日) 是假日，24號(六)也是假日，23號(五)是工作日
  const holidays = [
    new Date(2024, 7, 25).getTime(),
    new Date(2024, 7, 24).getTime()
  ];
  const workdays = [];
  const result = getSendDate(2024, 8, holidays, workdays);
  assertEquals('getSendDate: 連假往前推', new Date(2024, 7, 23).toISOString().slice(0,10), result.toISOString().slice(0,10));
}

function test_processEmailTemplates_variables() {
  // 測試變數替換
  const settings = {
    subjectNormal: '【通知】{{rocYear}}年{{currentMonth}}月款項申請(至{{deadlineDate}}前截止)',
    bodyNormal: '民國年:{{rocYear}}, 月份:{{currentMonth}}, 截止:{{deadlineDate}}',
    subjectDecember: '【通知】{{rocYear}}年{{currentMonth}}月款項申請，至{{deadlineDate}}截止。',
    bodyDecember: '民國年:{{rocYear}}, 月份:{{currentMonth}}, 明年:{{nextRocYear}}, 截止:{{deadlineDate}}'
  };
  const holidays = [];
  const workdays = [];
  // mock calculateDeadline
  const oldCalculateDeadline = this.calculateDeadline;
  this.calculateDeadline = function(year, month, h, w) { return '測試截止日'; };
  // 8月
  let result = processEmailTemplates(settings, 2024, 8);
  assertDeepEquals('processEmailTemplates: 8月', {subject:'【通知】113年8月款項申請(至測試截止日前截止)', body:'民國年:113, 月份:8, 截止:測試截止日'}, result);
  // 12月
  result = processEmailTemplates(settings, 2024, 12);
  assertDeepEquals('processEmailTemplates: 12月', {subject:'【通知】113年12月款項申請，至測試截止日截止。', body:'民國年:113, 月份:12, 明年:114, 截止:測試截止日'}, result);
  this.calculateDeadline = oldCalculateDeadline;
}
