const dotenv = require("dotenv");
dotenv.config();

const { Metro } = require('../classes/metro');
const { MerryMart } = require('../classes/merrymart');

const buildMetro = async(store, action) => {
    const metro = new Metro();
    metro.chain = store;
    metro.action = action;
    const { isProcessed, statusMsg } = await metro.buildRawData();
    if (isProcessed) {
        metro.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

const buildMerryMart = async(store, action) => {
    const merrymart = new MerryMart();
    merrymart.chain = store;
    merrymart.action = action;
    const { isProcessed, statusMsg } = await merrymart.buildRawData();
    if (isProcessed) {
        merrymart.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }  
}

module.exports = { buildMetro, buildMerryMart }