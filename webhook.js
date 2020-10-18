const EWS = require('node-ews');
const dotenv = require('dotenv');
const express = require("express")
const bodyParser = require("body-parser")
const EventEmitter = require('events');
dotenv.config();

class MyEmitter extends EventEmitter {}

const myEmitter = new MyEmitter();

// exchange server connection info
const ewsConfig = {
  username: process.env.USERNAME,
  password: process.env.PASSWORD,
  host: process.env.HOST
};

// initialize node-ews
const ews = new EWS(ewsConfig);

// define ews api function
const ewsFunction = 'CreateItem';

// Initialize express and define a port
const app = express()
const PORT = 3000

// Tell express to use body-parser's JSON parsing
app.use(bodyParser.json())

// Start express on the defined port
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`))

app.post("/hook", (req, res) => {
  console.log(req.headers)
  myEmitter.emit('sendemail', req.headers);
  res.status(200).end()
})


myEmitter.on('sendemail', function(requestBody) {
  console.log('someone sent a request!');
  console.log('sending email');

  // define ews api function args
  const ewsArgs = {
    "attributes" : {
      "MessageDisposition" : "SendAndSaveCopy"
    },
    "SavedItemFolderId": {
      "DistinguishedFolderId": {
        "attributes": {
          "Id": "sentitems"
        }
      }
    },
    "Items" : {
      "Message" : {
        "ItemClass": "IPM.Note",
        "Subject" : "Test EWS Email",
        "Body" : {
          "attributes": {
            "BodyType" : "Text"
          },
          "$value": requestBody.form_data
        },
        "ToRecipients" : {
          "Mailbox" : {
            "EmailAddress" : "gashackdummy@gmail.com"
          }
        },
        "IsRead": "false"
      }
    }
  };

  ews.run(ewsFunction, ewsArgs)
    .then(result => {
      console.log(JSON.stringify(result));
    })
    .catch(err => {
      console.log(err.stack);
    });
});
