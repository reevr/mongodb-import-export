const MongoClient = require('mongodb').MongoClient;

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

async function boot() {
  const connection = await connect('mongodb://localhost:27017')
  var mongoUtil = require('./mongo-util')(connection);

  mongoUtil.export('absentia', 'fbtargetings');
}

boot();