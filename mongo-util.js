var Excel = require('exceljs');
var path = require('path');

function mongoUtility(connection) {

  async function Import(dbName, collection, filePath) {
    try {
      var db = setDB(dbName);
      collection = setCollection(db, collection);
      var workbook = new Excel.Workbook();
      await workbook.xlsx.readFile(filePath)
      var worksheet = workbook.getWorksheet(1);
      const data = [];
      let columns;
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber == 1) {
          const [empty, ...values] = row.values;
          delete empty;
          columns = values;
        }
        data.push(getKeyValueMapping(columns, row.values))
      });
      await insertManyDocuments(data, collection)

    } catch (err) {
      throw err;
    }
  }

  function insertManyDocuments(data, collection) {
    return new Promise((resolve, reject) => {
      collection.insertMany(data, function (err, res) {
        if (err) throw err;
        console.log("Number of documents inserted: " + res.insertedCount);
      });
    });
  }

  function getKeyValueMapping(columns, [empty, ...values]) {
    const obj = {};
    values.forEach((value, index) => {
      if (columns[index] !== '_id') {
        obj[columns[index]] = value;
      }
    });
    return obj;
  }

  async function createOrUpdateGroup(dbName, collection, filePath) {
    try {
      var db = setDB(dbName);
      var workbook = new Excel.Workbook();
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1);

      const groupTargetingMap = groupByGroupName(worksheet);
      const _idIndex = worksheet.getRow(1).values.findIndex(col => col == '_id');
      const results = await Promise.all(syncWithDB(groupTargetingMap, db, _idIndex))
    } catch (err) {
      throw err;
    }
  }

  function syncWithDB(groupTargetingMap, db, _idIndex) {
    const collection = setCollection(db, 'groups');
    return Object.keys(groupTargetingMap).map(groupName => {
      return new Promise(async (resolve, reject) => {
        try {
          const result = await findDocument(collection, { group_name: groupName });
          let targetingIds = groupTargetingMap[groupName].map(value => value[_idIndex])
          if (!!result) {
            result.targeting_ids = [...result.targeting_ids, ...targetingIds];
            await upsertDocument({ group_name: groupName }, { targeting_ids: result.targeting_ids }, collection);
          } else {
            await upsertDocument({ group_name: groupName }, { targeting_ids: result.targetingIds }, collection);
          }
          resolve(true);
        } catch (err) {
          reject(err);
        }
      });
    })
  }

  function upsertDocument(searchQuery, newValues, collection) {
    return new Promise((resolve, reject) => {
      collection.update(
        searchQuery,
        newValues,
        { upsert: true, safe: false },
        function (err, data) {
          if (err) {
            return reject(err);
          }
          resolve(data);
        }
      );
    })
  }

  function findDocument(collection, query) {
    return new Promise((resolve, reject) => {
      collection.findOne(query, (err, result) => {
        if (err) return reject(err);
        resolve(result);
      })
    });
  }

  function groupByGroupName(worksheet) {
    let groupIndex;
    const group = {};
    worksheet.eachRow(function (row, rowNumber) {
      if (rowNumber == 1) {
        groupIndex = row.values.findIndex(column => column == 'group');
      } else {
        groupBy(row.values[groupIndex], row.values, group);
      }
    });
    return group;
  }

  function groupBy(keys, value, output) {
    keys = keys.split(",").map(val => val.trim());
    keys.forEach(key => {
      if (!output[key]) {
        output[key] = [];
      }
      output[key].push(value);
    })
  }

  function Export(dbName, collection) {
    var db = setDB(dbName);
    collection = setCollection(db, collection);
    return getData(collection, {})
      .then(data => createWorksheet(data))
      .then(() => console.info('Exported succesfully'))
      .catch(err => console.error(err));
  }

  function createWorksheet(data) {
    console.log(data);
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

      const filename = +(new Date) + ".xlsx"
      const filePath = path.join(__dirname, './output', filename)
      workbook.xlsx.writeFile(filePath)
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
    export: Export,
    createOrUpdateGroup: createOrUpdateGroup
  }
}
module.exports = mongoUtility