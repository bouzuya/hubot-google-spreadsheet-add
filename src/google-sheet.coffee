google = require 'googleapis'
{Promise} = require 'es6-promise'
{parseString} = require 'xml2js'

class GoogleSheet
  @baseUrl: 'https://spreadsheets.google.com/feeds'

  @visibilities:
    private: 'private'
    public: 'public'

  @projections:
    basic: 'basic'
    full: 'full'

  constructor: ({ @credentials, @key }) ->
    @client = null # initialize in @authorize()

  authorize: ->
    @_authorize(@credentials)
    .then (client) =>
      @client = client

  put: (row, col, value) ->
    worksheetId = null
    cellUrl = null
    @_getWorksheetId()
    .then (w) -> worksheetId = w
    .then =>
      cellUrl = @_getCellUrl
        key: @key
        worksheetId: worksheetId
        visibilities: GoogleSheet.visibilities.private
        projections: GoogleSheet.projections.full
        row: row
        col: col
      @_request(@client, {
        url: cellUrl
        method: 'GET'
        headers:
          'GData-Version': '3.0'
          'Content-Type': 'application/atom+xml'
      })
    .then @_parseXml.bind(@)
    .then (data) =>
      xml = """
        <entry xmlns="http://www.w3.org/2005/Atom"
            xmlns:gs="http://schemas.google.com/spreadsheets/2006">
          <id>#{cellUrl}</id>
          <link rel="edit" type="application/atom+xml" href="#{cellUrl}"/>
          <gs:cell row="#{row}" col="#{col}" inputValue="#{value}"/>
        </entry>
      """
      @_request(@client, {
        url: cellUrl
        method: 'PUT'
        headers:
          'GData-Version': '3.0'
          'Content-Type': 'application/atom+xml'
          'If-Match': data.entry.$['gd:etag']
        body: xml
      })
    .then @_parseXml.bind(@)

  getCells: (col) ->
    columnName = @_getColumnName(col)
    @_getWorksheetId()
    .then (worksheetId) =>
      cellsUrl = @_getCellsUrl
        key: @key
        worksheetId: worksheetId
        visibilities: GoogleSheet.visibilities.private
        projections: GoogleSheet.projections.basic
      @_request(@client, { url: cellsUrl })
    .then @_parseXml.bind(@)
    .then (data) ->
      data.feed.entry
      .map (i) ->
        { title: i.title[0]._, content: i.content[0]._ }
      .filter (i) ->
        i.title.match(new RegExp('^' + columnName))

  load: ->
    worksheetUrl = @_getWorksheetUrl
      key: @key
      visibilities: GoogleSheet.visibilities.private
      projections: GoogleSheet.projections.basic

    @_authorize(@credentials)
    .then => @_request(@client, { url: worksheetUrl })
    .then @_parseXml.bind(@)
    .then (data) ->
      worksheetUrls = data.feed.entry.map (i) -> i.id[0]
      url = worksheetUrls[0]
      throw new Error() if url.indexOf(worksheetUrl) isnt 0
      url.replace(worksheetUrl + '/', '')
    .then (worksheetId) =>
      cellsUrl = @_getCellsUrl
        key: @key
        worksheetId: worksheetId
        visibilities: GoogleSheet.visibilities.private
        projections: GoogleSheet.projections.basic
      @_request(@client, { url: cellsUrl })
    .then @_parseXml.bind(@)
    .then (data) ->
      data.feed.entry.map (i) ->
        { title: i.title[0]._, content: i.content[0]._ }

  _authorize: ({ email, key })->
    new Promise (resolve, reject) ->
      scope = ['https://spreadsheets.google.com/feeds']
      jwt = new google.auth.JWT(email, null, key, scope, null)
      jwt.authorize (err) ->
        if err? then reject(err) else resolve(jwt)

  _getCell: (row, col) ->
    @_getWorksheetId()
    .then (worksheetId) =>
      cellUrl = @_getCellUrl
        key: @key
        worksheetId: worksheetId
        visibilities: GoogleSheet.visibilities.private
        projections: GoogleSheet.projections.full
        row: row
        col: col
      @_request(@client, {
        url: cellUrl
        method: 'GET'
        headers:
          'GData-Version': '3.0'
          'Content-Type': 'application/atom+xml'
      })
    .then @_parseXml.bind(@)

  # visibilities: private / public
  # projections: full / basic
  _getCellUrl: ({ key, worksheetId, visibilities, projections, row, col }) ->
    path = "/cells/#{key}/#{worksheetId}/#{visibilities}/#{projections}/R#{row}C#{col}"
    GoogleSheet.baseUrl + path

  # visibilities: private / public
  # projections: full / basic
  _getCellsUrl: ({ key, worksheetId, visibilities, projections }) ->
    path = "/cells/#{key}/#{worksheetId}/#{visibilities}/#{projections}"
    GoogleSheet.baseUrl + path

  _getColumnName: (col) ->
    String.fromCharCode('A'.charCodeAt(0) + col - 1)

  # visibilities: private / public
  # projections: full / basic
  _getWorksheetUrl: ({ key, visibilities, projections }) ->
    path = "/worksheets/#{key}/#{visibilities}/#{projections}"
    GoogleSheet.baseUrl + path

  _getWorksheetId: ->
    worksheetUrl = @_getWorksheetUrl
      key: @key
      visibilities: GoogleSheet.visibilities.private
      projections: GoogleSheet.projections.basic

    @_request(@client, { url: worksheetUrl })
    .then @_parseXml.bind(@)
    .then (data) ->
      worksheetUrls = data.feed.entry.map (i) -> i.id[0]
      url = worksheetUrls[0]
      throw new Error() if url.indexOf(worksheetUrl) isnt 0
      url.replace(worksheetUrl + '/', '')

  _listSpreadsheetsUrl: (client) ->
    visibilities = GoogleSheet.visibilities.private
    projections = GoogleSheet.projections.full
    path = "/spreadsheets/#{visibilities}/#{projections}"
    GoogleSheet.baseUrl + path

  _parseXml: (xml) ->
    new Promise (resolve, reject) ->
      parseString xml, (err, parsed) ->
        if err? then reject(err) else resolve(parsed)

  _request: (client, options) ->
    new Promise (resolve, reject) ->
      client.request options, (err, data) ->
        if err? then reject(err) else resolve(data)

module.exports.GoogleSheet = GoogleSheet
