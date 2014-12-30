
var GoogleSheet, Promise, google, parseString;

google = require('googleapis');

Promise = require('es6-promise').Promise;

parseString = require('xml2js').parseString;

GoogleSheet = (function() {
  GoogleSheet.baseUrl = 'https://spreadsheets.google.com/feeds';

  GoogleSheet.visibilities = {
    "private": 'private',
    "public": 'public'
  };

  GoogleSheet.projections = {
    basic: 'basic',
    full: 'full'
  };

  function GoogleSheet(_arg) {
    this.credentials = _arg.credentials, this.key = _arg.key;
    this.client = null;
  }

  GoogleSheet.prototype.authorize = function() {
    return this._authorize(this.credentials).then((function(_this) {
      return function(client) {
        return _this.client = client;
      };
    })(this));
  };

  GoogleSheet.prototype.put = function(row, col, value) {
    var cellUrl, worksheetId;
    worksheetId = null;
    cellUrl = null;
    return this._getWorksheetId().then(function(w) {
      return worksheetId = w;
    }).then((function(_this) {
      return function() {
        cellUrl = _this._getCellUrl({
          key: _this.key,
          worksheetId: worksheetId,
          visibilities: GoogleSheet.visibilities["private"],
          projections: GoogleSheet.projections.full,
          row: row,
          col: col
        });
        return _this._request(_this.client, {
          url: cellUrl,
          method: 'GET',
          headers: {
            'GData-Version': '3.0',
            'Content-Type': 'application/atom+xml'
          }
        });
      };
    })(this)).then(this._parseXml.bind(this)).then((function(_this) {
      return function(data) {
        var xml;
        xml = "<entry xmlns=\"http://www.w3.org/2005/Atom\"\n    xmlns:gs=\"http://schemas.google.com/spreadsheets/2006\">\n  <id>" + cellUrl + "</id>\n  <link rel=\"edit\" type=\"application/atom+xml\" href=\"" + cellUrl + "\"/>\n  <gs:cell row=\"" + row + "\" col=\"" + col + "\" inputValue=\"" + value + "\"/>\n</entry>";
        return _this._request(_this.client, {
          url: cellUrl,
          method: 'PUT',
          headers: {
            'GData-Version': '3.0',
            'Content-Type': 'application/atom+xml',
            'If-Match': data.entry.$['gd:etag']
          },
          body: xml
        });
      };
    })(this)).then(this._parseXml.bind(this));
  };

  GoogleSheet.prototype.getCells = function(col) {
    var columnName;
    columnName = this._getColumnName(col);
    return this._getWorksheetId().then((function(_this) {
      return function(worksheetId) {
        var cellsUrl;
        cellsUrl = _this._getCellsUrl({
          key: _this.key,
          worksheetId: worksheetId,
          visibilities: GoogleSheet.visibilities["private"],
          projections: GoogleSheet.projections.basic
        });
        return _this._request(_this.client, {
          url: cellsUrl
        });
      };
    })(this)).then(this._parseXml.bind(this)).then(function(data) {
      return data.feed.entry.map(function(i) {
        return {
          title: i.title[0]._,
          content: i.content[0]._
        };
      }).filter(function(i) {
        return i.title.match(new RegExp('^' + columnName));
      });
    });
  };

  GoogleSheet.prototype.load = function() {
    var worksheetUrl;
    worksheetUrl = this._getWorksheetUrl({
      key: this.key,
      visibilities: GoogleSheet.visibilities["private"],
      projections: GoogleSheet.projections.basic
    });
    return this._authorize(this.credentials).then((function(_this) {
      return function() {
        return _this._request(_this.client, {
          url: worksheetUrl
        });
      };
    })(this)).then(this._parseXml.bind(this)).then(function(data) {
      var url, worksheetUrls;
      worksheetUrls = data.feed.entry.map(function(i) {
        return i.id[0];
      });
      url = worksheetUrls[0];
      if (url.indexOf(worksheetUrl) !== 0) {
        throw new Error();
      }
      return url.replace(worksheetUrl + '/', '');
    }).then((function(_this) {
      return function(worksheetId) {
        var cellsUrl;
        cellsUrl = _this._getCellsUrl({
          key: _this.key,
          worksheetId: worksheetId,
          visibilities: GoogleSheet.visibilities["private"],
          projections: GoogleSheet.projections.basic
        });
        return _this._request(_this.client, {
          url: cellsUrl
        });
      };
    })(this)).then(this._parseXml.bind(this)).then(function(data) {
      return data.feed.entry.map(function(i) {
        return {
          title: i.title[0]._,
          content: i.content[0]._
        };
      });
    });
  };

  GoogleSheet.prototype._authorize = function(_arg) {
    var email, key;
    email = _arg.email, key = _arg.key;
    return new Promise(function(resolve, reject) {
      var jwt, scope;
      scope = ['https://spreadsheets.google.com/feeds'];
      jwt = new google.auth.JWT(email, null, key, scope, null);
      return jwt.authorize(function(err) {
        if (err != null) {
          return reject(err);
        } else {
          return resolve(jwt);
        }
      });
    });
  };

  GoogleSheet.prototype._getCell = function(row, col) {
    return this._getWorksheetId().then((function(_this) {
      return function(worksheetId) {
        var cellUrl;
        cellUrl = _this._getCellUrl({
          key: _this.key,
          worksheetId: worksheetId,
          visibilities: GoogleSheet.visibilities["private"],
          projections: GoogleSheet.projections.full,
          row: row,
          col: col
        });
        return _this._request(_this.client, {
          url: cellUrl,
          method: 'GET',
          headers: {
            'GData-Version': '3.0',
            'Content-Type': 'application/atom+xml'
          }
        });
      };
    })(this)).then(this._parseXml.bind(this));
  };

  GoogleSheet.prototype._getCellUrl = function(_arg) {
    var col, key, path, projections, row, visibilities, worksheetId;
    key = _arg.key, worksheetId = _arg.worksheetId, visibilities = _arg.visibilities, projections = _arg.projections, row = _arg.row, col = _arg.col;
    path = "/cells/" + key + "/" + worksheetId + "/" + visibilities + "/" + projections + "/R" + row + "C" + col;
    return GoogleSheet.baseUrl + path;
  };

  GoogleSheet.prototype._getCellsUrl = function(_arg) {
    var key, path, projections, visibilities, worksheetId;
    key = _arg.key, worksheetId = _arg.worksheetId, visibilities = _arg.visibilities, projections = _arg.projections;
    path = "/cells/" + key + "/" + worksheetId + "/" + visibilities + "/" + projections;
    return GoogleSheet.baseUrl + path;
  };

  GoogleSheet.prototype._getColumnName = function(col) {
    return String.fromCharCode('A'.charCodeAt(0) + col - 1);
  };

  GoogleSheet.prototype._getWorksheetUrl = function(_arg) {
    var key, path, projections, visibilities;
    key = _arg.key, visibilities = _arg.visibilities, projections = _arg.projections;
    path = "/worksheets/" + key + "/" + visibilities + "/" + projections;
    return GoogleSheet.baseUrl + path;
  };

  GoogleSheet.prototype._getWorksheetId = function() {
    var worksheetUrl;
    worksheetUrl = this._getWorksheetUrl({
      key: this.key,
      visibilities: GoogleSheet.visibilities["private"],
      projections: GoogleSheet.projections.basic
    });
    return this._request(this.client, {
      url: worksheetUrl
    }).then(this._parseXml.bind(this)).then(function(data) {
      var url, worksheetUrls;
      worksheetUrls = data.feed.entry.map(function(i) {
        return i.id[0];
      });
      url = worksheetUrls[0];
      if (url.indexOf(worksheetUrl) !== 0) {
        throw new Error();
      }
      return url.replace(worksheetUrl + '/', '');
    });
  };

  GoogleSheet.prototype._listSpreadsheetsUrl = function(client) {
    var path, projections, visibilities;
    visibilities = GoogleSheet.visibilities["private"];
    projections = GoogleSheet.projections.full;
    path = "/spreadsheets/" + visibilities + "/" + projections;
    return GoogleSheet.baseUrl + path;
  };

  GoogleSheet.prototype._parseXml = function(xml) {
    return new Promise(function(resolve, reject) {
      return parseString(xml, function(err, parsed) {
        if (err != null) {
          return reject(err);
        } else {
          return resolve(parsed);
        }
      });
    });
  };

  GoogleSheet.prototype._request = function(client, options) {
    return new Promise(function(resolve, reject) {
      return client.request(options, function(err, data) {
        if (err != null) {
          return reject(err);
        } else {
          return resolve(data);
        }
      });
    });
  };

  return GoogleSheet;

})();

module.exports.GoogleSheet = GoogleSheet;
