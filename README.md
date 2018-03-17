# Explainable Formulas

Making Google Sheets formulas readable by humans

Made by Glen Chiacchieri + Geoffrey Litt

## Development

Install Google's [clasp](https://github.com/google/clasp) tool for local Google Apps Script development:

`npm i @google/clasp -g`

`clasp` has strict rules about what types of files can be in the clasp project,
so for now all code is in the `src` directory and we need to run all `clasp`
commands within that directory (In theory `.claspignore` is supposed to deal
with this but it doesn't seem to work right now):

`cd src`

Clone the Google Apps Script development project:

`clasp clone 1k9IcK5YVvgVO6XiEJ98gJzBOBexas2k4fRKPXlt8schAOu2HrW8V_z-7`

Make changes locally.

Push changes to Google Apps Script project:

`clasp push`

In the test spreadsheet, reload the extension to run new code by clicking `Add-ons > Explainable Formulas 2 > Open Sidebar`

----


