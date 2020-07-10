var express = require('express');
var app = express();
const CosmosClient = require("@azure/cosmos").CosmosClient;
const schema = require('./schema')

const config = require("./config");
const url = require("url");
const readXlsxFile = require('read-excel-file/node');

const endpoint = config.endpoint;
const key = config.key;

const databaseId = config.database.id;
const containerId = config.container.id;
const partitionKey = { kind: "Hash", paths: ["/isFree"] };
const https = require("https");
const { items } = require("./config");
const { promises } = require("fs");

const options = {
    endpoint: endpoint,
    key: key,
    userAgentSuffix: "CosmosDBJavascriptQuickstart",
    agent: new https.Agent({
        rejectUnauthorized: false,
    }),
};

const client = new CosmosClient(options);

async function excelToJson() {
    const file = __dirname + '\\Book2.xlsx';

    let { rows, error } = await readXlsxFile(file, { schema });
    // const marvel = excel.SheetNames;
    // const test = await json(excel.Sheets[marvel[0]])
    await addToDb(rows);
};

excelToJson();



async function addToDb(datas) {
    const container = client.database(databaseId).container(containerId)
    const creates = []
    for (const data of datas) {
        creates.push(container.items.create(data))
    }
    return Promise.all(creates)
}


