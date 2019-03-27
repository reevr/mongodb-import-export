const MongoClient = require('mongodb').MongoClient;
const path = require('path')

function connect(url) {
  return new Promise((resolve, reject) => {
    MongoClient.connect(url, function (err, client) {
      if (err) {
        console.error(err);
        return reject(err)
      }
      console.log("Connected successfully to server");
      resolve(client);
    });
  })
}

async function boot(type, dbName, collection, filePath, url = 'mongodb://localhost:27017') {
  const connection = await connect(url)
  var mongoUtil = require('./mongo-util')(connection);

  if (type === 'export') {
    await mongoUtil.export(dbName, collection);
  } else if (type == 'import') {
    await mongoUtil.import(dbName, collection, filePath);
  } else {
    await mongoUtil.createOrUpdateGroup(dbName, collection, filePath);
  }
}

const type = process.argv[2] || 'export';
const dbName = process.argv[3] || 'absentia';
const collection = process.argv[4] || 'fbtargetings';
const filePath = process.argv[4] ? path.join(__dirname, './output', process.argv[4]) : path.join(__dirname, './output', 'mongodb-export.xlsx');

boot(type, dbName, collection, filePath).then(() => console.log('done'));


console.log('******', type)
