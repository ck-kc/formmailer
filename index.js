const EWS = require('node-ews');
const dotenv = require('dotenv');
dotenv.config();

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
        "$value": "This is a test email"
      },
      "ToRecipients" : {
        "Mailbox" : {
          "EmailAddress" : "c.razavi@nhs.net"
        }
      },
      "IsRead": "false"
    }
  }
};

// query ews, print resulting JSON to console
ews.run(ewsFunction, ewsArgs)
  .then(result => {
    console.log(JSON.stringify(result));
  })
  .catch(err => {
    console.log(err.stack);
  });
