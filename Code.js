/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Runs when the add-on is installed.
 */
function onInstall() {
  onOpen();
}


/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Open Sidebar', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Explain Formula');
  SpreadsheetApp.getUi().showSidebar(ui);
}

// Hack to evaluate a formula, because there's no built-in eval.
// We just copy the formula to a random cell and get the result from there.
function evalFormula(formula) {
  // Obviously in practice would need to choose the scratch cell more carefully than this.
  var scratchCell = SpreadsheetApp.getActiveSheet().getRange(20, 10);
  scratchCell.setFormula(formula);
  return scratchCell.getValue();
}

// Get sub-expressions of the formula and their evaluated results.
//
// For the moment, this is just a super naive proof of concept.
// All it does is extract leaf expressions with no nested parens.
// 
// for real subexpression parsing a better approach is needed.
// This JS Excel formula parser might be useful:
// http://ewbi.blogs.com/develops/2004/12/excel_formula_p.html
function getSubExpressions(formula) {
  var subExpressions = formula.match(/[A-Z]*\([^\(\)]*\)/g);
  var subExpressionsWithResults = []
  
  if (subExpressions !== null) {
    subExpressions.forEach(function (subexpr) {
      subExpressionsWithResults.push(
        {
          formula: subexpr,
          result: evalFormula(subexpr)
        }
      );
    });
  }
  
  return subExpressionsWithResults;
}

// Get the current formula and its subexpressions w/ results
function getCurrentFormula() {
  var formula = SpreadsheetApp.getActiveRange().getFormula();
  
  return {
    formula: formula,
    subexpressions: getSubExpressions(formula)
  }
}

/**
 * Find a cell in the spreadsheet that isn't referenced
 */
function findTestCell() {
  SpreadsheetApp
}

