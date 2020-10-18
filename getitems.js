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
const ewsFunction = 'FindItem';

// define custom soap header
const ewsSoapHeader = {
  't:RequestServerVersion': {
    attributes: {
      Version: "Exchange2013_SP1"
    }
  }
};

// define ews api function args
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
  "QueryString" : "subject:Test EWS Email"
};

// query ews, print resulting JSON to console
ews.run(ewsFunction, ewsArgs, ewsSoapHeader)
  .then(result => {
    console.log(JSON.stringify(result));
  })
  .catch(err => {
    console.log(err.stack);
  });
