# Description
#   A Hubot script to add the value to google spreadsheet cell
#
# Configuration:
#   None
#
# Commands:
#   hubot google spreadsheet add <value> - add the value to google spreadsheet cell
#
# Author:
#   bouzuya <m@bouzuya.net>
#
{GoogleSheet} = require '../google-sheet'
parseConfig = require 'hubot-config'

config = parseConfig 'auto-lgtm',
  googleEmail: null
  googleKey: null
  googleSheetKey: null

module.exports = (robot) ->
  column = 1
  robot.respond /google spreadsheet add (.+)$/i, (res) ->
    row = null
    message = res.match[1]
    sheet = new GoogleSheet
      credentials:
        email: config.googleEmail
        key: config.googleKey
      key: config.googleSheetKey
    sheet.authorize()
    .then ->
      sheet.getCells(column)
    .then (cells) ->
      cells[cells.length - 1]
    .then (cell) ->
      row = parseInt(cell.title.match(/^[A-Z]+(\d+)$/)[1], 10)
      sheet.put(row + 1, column, message)
    .then ->
      res.send "R#{row + 1}C#{column} <- #{message}"
