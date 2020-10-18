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

// define custom soap header
const ewsSoapHeader = {
  't:RequestServerVersion': {
    attributes: {
      Version: "Exchange2013_SP1"
    }
  }
};

// initialize node-ews
const ews = new EWS(ewsConfig);



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


async function getItems(searchResult) {

  const ewsFunction = 'GetItem';

  let messageArray = searchResult.ResponseMessages.FindItemResponseMessage.RootFolder.Items.Message;
  console.log(messageArray);

  let idList = messageArray.map(x => x.ItemId.attributes.Id);
  console.log(idList);

  const ewsArgs = {
    "ItemShape": {
      "BaseShape" : "Default",
    },
    "ItemIds": {
      idList
    },
    "ParentFolderIds": {
      "DistinguishedFolderId": {
        "attributes": {
          "Id": "inbox"
        }
      }
    },
  };

  ews.run(ewsFunction, ewsArgs, ewsSoapHeader)
    .then(result => {
      console.log(JSON.stringify(result));


      res.send(JSON.stringify(result));
      res.status(200).end();
    })
    .catch(err => {
      console.log(err.stack);
      res.status(500).end();
    });

}

app.get("/getitems2", (req, res) => {

  // define ews api function
  const ewsFunction = 'FindItem';

  const ewsArgs = {
    "attributes" : {
      "Traversal" : "Shallow"
    },
    "ItemShape": {
      "BaseShape" : "IdOnly",
      "AdditionalProperties": {
        "FieldURI": {
          "attributes" : {
            "FieldURI": "item:Subject",
          }
        }
      }
    },
    "ParentFolderIds": {
      "DistinguishedFolderId": {
        "attributes": {
          "Id": "inbox"
        }
      }
    },
    "QueryString" : "subject:FormMailerRecord"
  };

  console.log('Someone is requesting items');
  console.log(req.headers);

  ews.run(ewsFunction, ewsArgs, ewsSoapHeader)
    .then(result => {
      console.log(JSON.stringify(result));
      getItems(result);

      res.send(JSON.stringify(result));
      res.status(200).end();
    })
    .catch(err => {
      console.log(err.stack);
      res.status(500).end();
    });

})


app.get("/getitems", (req, res) => {

  // define ews api function
  const ewsFunction = 'FindItem';

  const ewsArgs = {
    "attributes" : {
      "Traversal" : "Shallow"
    },
    "ItemShape": {
      "BaseShape" : "IdOnly",
      "AdditionalProperties": {
        "FieldURI": {
          "attributes" : {
            "FieldURI": "item:Subject",
          }
        }
      }
    },
    "ParentFolderIds": {
      "DistinguishedFolderId": {
        "attributes": {
          "Id": "inbox"
        }
      }
    },
    "QueryString" : "subject:FormMailerRecord"
  };

  console.log('Someone is requesting items');
  console.log(req.headers);

  ews.run(ewsFunction, ewsArgs, ewsSoapHeader)
    .then(result => {
      console.log(JSON.stringify(result));

      res.send(JSON.stringify(result));
      res.status(200).end();
    })
    .catch(err => {
      console.log(err.stack);
      res.status(500).end();
    });

})


myEmitter.on('sendemail', function(requestBody) {
  console.log('someone sent a request!');
  console.log('sending email');

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
        "Subject" : "FormMailerRecord " + requestBody.form_data,
        "Body" : {
          "attributes": {
            "BodyType" : "Text"
          },
          "$value": requestBody.form_data
        },
        "ToRecipients" : {
          "Mailbox" : {
            "EmailAddress" : process.env.USERNAME
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
