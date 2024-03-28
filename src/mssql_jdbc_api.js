
//Michal Jakubik fecit, Georg Wuitschik modifecit.
class mssql_jdbc_api {

  constructor(server, port, dbName, username, password) {
    this.server = server;
    this.port = port;
    this.dbName = dbName;
    this.username = username;
    this.password = password;

    this.url = 'jdbc:sqlserver://' + this.server + ':' + this.port + ';databaseName=' + this.dbName;

    this.connect();

    this.setBatchSize(100);
  }

  setBatchSize(batchsize) {
    this.batchSize = batchsize;
  }

  connect() {
    this.conn = Jdbc.getConnection(this.url, this.username, this.password);
  }

  disconnect() {
    this.conn.close();
  }

  getResults() {
    return this.results;
  }
  // execute is used for queries that don't return anything, e.g. update, delete from table where ...
  execute(query) {
    var stmt = this.conn.createStatement();
    stmt.execute(query);
  }
  //'SELECT * FROM dbo.msdata' , used for queries that return something, e.g. select,  insert
  executeQuery(query) {
    var stmt = this.conn.createStatement();
    this.results = stmt.executeQuery(query);
  }

  getResultsAsArray() {
    var resultsArray = [];
    var metaData = this.results.getMetaData();
    this.numCols = metaData.getColumnCount();

    while (true) {
      var batcharr = this.generateRowItemsBatch(this.batchSize);
      if (batcharr.length == 0) break;
      //arr.push(counter);
      //sheet.appendRow(arr);
      resultsArray = resultsArray.concat(batcharr);
    }
    this.results.close(); // used to be after the return statement, where it can't be reached. 
    return resultsArray;
  }

  pushResultsToSheet(sheetName) {
    var resultsArray = [];
    var metaData = this.results.getMetaData();
    this.numCols = metaData.getColumnCount();
    var spreadsheet = SpreadsheetApp.getActive();
    var sheet = spreadsheet.getSheetByName(sheetName);
    sheet.clearContents();
    var arr = [];

    for (var col = 0; col < this.numCols; col++) {
      arr.push(metaData.getColumnName(col + 1));
    }

    sheet.appendRow(arr);
    var counter = 0;
    while (true) {
      var batcharr = this.generateRowItemsBatch(this.batchSize);
      if (batcharr.length == 0) break;
      var lastRow = sheet.getLastRow();
      //arr.push(counter);
      //sheet.appendRow(arr);
      sheet.getRange(lastRow + 1, 1, batcharr.length, batcharr[0].length).setValues(batcharr);
    }

    this.results.close();
    //stmt.close();
    //sheet.autoResizeColumns(1, numCols+1);
  }

  generateRowItemsBatch(batchsize) {
    var resarr = [];
    var counter = 0;
    while (this.results.next()) {
      counter++;
      var rowarr = [];
      for (var col = 0; col < this.numCols; col++) {
        rowarr.push(this.results.getString(col + 1));
      }
      resarr.push(rowarr);
      if (counter >= batchsize) break;
    }

    return resarr;
  }

  upsertData(statement) {
    var stmt = this.conn.createStatement();
    this.results = stmt.executeUpdate(statement);

    if (this.results > -1) return "Execution OK. " + this.results + " rows updated";
    return "Error";
  }

  updateDataDictionary(table, datadict) {
    this.conn.setAutoCommit(false);
    var stmt = this.conn.createStatement();
    var statement = "";
    for (const [key, value] of Object.entries(datadict)) {
      var ID = "";
      var DATA = "";
      statement = "update " + table + " set ";
      for (const [upskey, upsvalue] of Object.entries(value)) {
        if (upskey == "ID") {
          for (const [upsidkey, upsidvalue] of Object.entries(upsvalue)) {
            ID += " " + upsidkey + "='" + upsidvalue + "' AND ";
          }
        }
        if (upskey == "DATA") {
          for (const [upsdatakey, upsdatavalue] of Object.entries(upsvalue)) {
            if (upsdatavalue == "NULL") {
              DATA += " " + upsdatakey + "= NULL, ";
            } else {
              DATA += " " + upsdatakey + "='" + upsdatavalue + "', ";
            }
          }
        }
      }
      DATA = DATA.slice(0, -2);
      ID = ID.slice(0, -5);
      statement += DATA + " where " + ID + ";";
      stmt.addBatch(statement);
    }
    console.log(statement);
    this.results = stmt.executeBatch();
    this.conn.commit();
    return "Execution OK";
  }

  insertDataDictionary(table, datadict) {
    this.conn.setAutoCommit(false);
    var stmt = this.conn.createStatement();
    var statement = "";
    for (const [key, value] of Object.entries(datadict)) {
      var ID = "";
      var DATA = "";
      statement = "insert into " + table + " ";
      for (const [upskey, upsvalue] of Object.entries(value)) {
        if (upskey == "ID") {
          for (const [upsidkey, upsidvalue] of Object.entries(upsvalue)) {
            DATA += ",'" + upsidvalue + "'";
            ID += "," + upsidkey;
          }
        }
        if (upskey == "DATA") {
          for (const [upsdatakey, upsdatavalue] of Object.entries(upsvalue)) {
            if (upsdatavalue == "NULL") {
              DATA += ", NULL";
            } else {
              DATA += ",'" + upsdatavalue + "'";
            }
            ID += "," + upsdatakey;
          }
        }
      }
      statement += "(" + ID.substring(1) + ") values (" + DATA.substring(1) + ");";
      stmt.addBatch(statement);
    }
    this.results = stmt.executeBatch();
    this.conn.commit();
    return "Execution OK";
  }
}
