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

// uses an Ohm grammar to parse and evaluate arithmetic
function evalArithmetic(value) {
  var g = ohm.grammar('Arithmetic {' + "\n" +
    '  Exp = AddExp' + "\n" +
    '  AddExp = AddExp "+" PriExp  -- plus' + "\n" +
    '         | AddExp "-" PriExp  -- minus' + "\n" +
    '         | PriExp' + "\n" +
    '  PriExp = "(" Exp ")"  -- paren' + "\n" +
    '         | number' + "\n" +
    '  number = digit+' + "\n" +
    '}')

  // Define an operation named 'eval' which evaluates the expression.
  // See https://github.com/cdglabs/ohm/blob/master/doc/api-reference.md#semantics
  var semantics = g.createSemantics().addOperation('eval', {
    Exp: function(e) {
      return e.eval();
    },
    AddExp: function(e) {
      return e.eval();
    },
    AddExp_plus: function(left, op, right) {
      return left.eval() + right.eval();
    },
    AddExp_minus: function(left, op, right) {
      return left.eval() - right.eval();
    },
    PriExp: function(e) {
      return e.eval();
    },
    PriExp_paren: function(open, exp, close) {
      return exp.eval();
    },
    number: function(chars) {
      return parseInt(this.sourceString, 10);
    },
  });

  var result;
  var m = g.match(value);
  if (m.succeeded()) {
    result = semantics(m).eval();  // Evaluate the expression.
  } else {
    result = m.message;  // Extract the error message.
  }

  return result;
}

function processCurrentCell () {
  var formula = SpreadsheetApp.getActiveRange().getFormula();
  var value = SpreadsheetApp.getActiveRange().getValue();

  console.log("processing", formula, value);

  return {
    formula: formula,
    subexpressions: getSubExpressions(formula),
    arithmeticResult: evalArithmetic(value)
  }
}

/**
 * Find a cell in the spreadsheet that isn't referenced
 */
function findTestCell() {
  SpreadsheetApp
}

