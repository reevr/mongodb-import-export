var Excel = require('exceljs');
var path = require('path');

function functions(connection) {

  function Import(bName, collection) {
    var workbook = new Excel.Workbook();
    var worksheet = workbook.addWorksheet('import-data');

  }

  function Export(dbName, collection) {
    var db = setDB(dbName);
    collection = setCollection(db, collection);
    getData(collection, {})
      .then(data => createWorksheet(data))
      .then(() => console.info('Exported succesfully'))
      .catch(err => console.error(err));
  }

  function createWorksheet(data) {
    return new Promise((resolve, reject) => {
      var workbook = new Excel.Workbook();
      var worksheet = workbook.addWorksheet('export-data');
      worksheet.columns = [
        { header: '_id', key: '_id', width: 10 },
        { header: '__v', key: '__v', width: 32 },
        { header: 'audience_size', key: 'audience_size', width: 10, outlineLevel: 1 },
        { header: 'city', key: 'city', width: 10, outlineLevel: 1 },
        { header: 'city_id', key: 'city_id', width: 10, outlineLevel: 1 },
        { header: 'common_name', key: 'common_name', width: 10, outlineLevel: 1 },
        { header: 'country_code', key: 'country_code', width: 10, outlineLevel: 1 },
        { header: 'country_name', key: 'country_name', width: 10, outlineLevel: 1 },
        { header: 'db_type', key: 'db_type', width: 10, outlineLevel: 1 },
        { header: 'field_id', key: 'field_id', width: 10, outlineLevel: 1 },
        { header: 'field_name', key: 'field_name', width: 10, outlineLevel: 1 },
        { header: 'id', key: 'id', width: 10, outlineLevel: 1 },
        { header: 'name', key: 'name', width: 10, outlineLevel: 1 },
        { header: 'path', key: 'path', width: 10, outlineLevel: 1 },
        { header: 'region', key: 'region', width: 10, outlineLevel: 1 },
        { header: 'region_id', key: 'region_id', width: 10, outlineLevel: 1 },
        { header: 'section_id', key: 'section_id', width: 10, outlineLevel: 1 },
        { header: 'section_name', key: 'section_name', width: 10, outlineLevel: 1 },
        { header: 'type', key: 'type', width: 10, outlineLevel: 1 },
        { header: 'unique_id', key: 'unique_id', width: 10, outlineLevel: 1 }
      ];

      if (data.length > 0) {
        for (let row of data) {
          worksheet.addRow(row);
        }
      }
      const filename = (new Date).toISOString() + ".xlsx"
      workbook.xlsx.writeFile(path.join(__dirname, './output', filename))
        .then(function () {
          resolve(true);
        }).catch(err => reject(err));
    })
  }

  function getData(collection, query) {
    return new Promise((resolve, reject) => {
      collection.find(query).toArray(function (err, docs) {
        if (err) {
          return reject(err)
        }
        resolve(docs);
      });
    })
  }

  function setDB(dbName) {
    return connection.db(dbName);
  }

  function setCollection(db, collection) {
    return db.collection(collection);
  }

  return {
    import: Import,
    export: Export
  }
}







module.exports = functions