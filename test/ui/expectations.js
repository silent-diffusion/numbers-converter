// What the UI smoke test (smoke.test.js) checks in particular files from test/fixtures, by their path
// there. Every .numbers file in that folder and the folders under it is opened, previewed sheet by
// sheet and saved whether it's listed here or not, so to test new files (say, ones saved by a newer
// Numbers), drop them in a folder there. List a file here only to check more of it:
//
//   sheets     sheet tabs in its preview (with Separate sheets, the default)
//   charts     charts drawn over its sheets, all sheets together
//   drawings   text boxes, shapes and pictures drawn over its sheets; pictures: how many are pictures
//   changes    what "What changes in Excel" lists to check, the number on the button beside Convert & save
//   oneSheet   { [Numbers sheet]: sheet tabs once that sheet is switched to One sheet }
//   password   for a password-protected file, its password; hint: the hint it shows
//   unreadable true for a file the app can't read yet: it must show the picture Numbers saved in it
module.exports = {
  'create-formulas.numbers': { sheets: 14, changes: 1, oneSheet: { 'Main Sheet': 6 } },
  'encrypted.numbers': { password: 's3cr3t', hint: 's3cr3t' },
  'folder-package.numbers': { sheets: 3 },
  'images.numbers': { sheets: 1, drawings: 6, pictures: 4, changes: 1 },
  'issue-66-collab.numbers': { sheets: 1, charts: 1 },
  'issue-9.numbers': { sheets: 7, charts: 10, drawings: 7, changes: 1 },
  'test-categories.numbers': { sheets: 11 },
  'test-custom-formats.numbers': { sheets: 2, changes: 2 },
  'test-pivot.numbers': { sheets: 6, oneSheet: { Sales: 4 } },
  'test-styles.numbers': { sheets: 4 }
};
