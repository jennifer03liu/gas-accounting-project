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
