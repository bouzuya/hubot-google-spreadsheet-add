// Description
//   A Hubot script to add the value to google spreadsheet cell
//
// Configuration:
//   None
//
// Commands:
//   hubot google spreadsheet add <value> - add the value to google spreadsheet cell
//
// Author:
//   bouzuya <m@bouzuya.net>
//
var GoogleSheet, config, parseConfig;

GoogleSheet = require('../google-sheet').GoogleSheet;

parseConfig = require('hubot-config');

config = parseConfig('auto-lgtm', {
  googleEmail: null,
  googleKey: null,
  googleSheetKey: null
});

module.exports = function(robot) {
  var column;
  column = 1;
  return robot.respond(/google spreadsheet add (.+)$/i, function(res) {
    var message, row, sheet;
    row = null;
    message = res.match[1];
    sheet = new GoogleSheet({
      credentials: {
        email: config.googleEmail,
        key: config.googleKey
      },
      key: config.googleSheetKey
    });
    return sheet.authorize().then(function() {
      return sheet.getCells(column);
    }).then(function(cells) {
      return cells[cells.length - 1];
    }).then(function(cell) {
      row = parseInt(cell.title.match(/^[A-Z]+(\d+)$/)[1], 10);
      return sheet.put(row + 1, column, message);
    }).then(function() {
      return res.send("R" + (row + 1) + "C" + column + " <- " + message);
    });
  });
};
