const { Robinson } = require('../classes/robinson');
const { Metro } = require('../classes/metro');
const { Puregold } = require('../classes/puregold');
const { MerryMart } = require('../classes/merrymart');
const { WalterMart } = require('../classes/waltermart');
const { WeShop } = require('../classes/weshop');

const consolidateRobinson = async(store, action, cutOff, salesType) => {
    const robinson = new Robinson();
    robinson.chain = store;
    robinson.action = action;
    robinson.cutOff = cutOff;
    robinson.salesType = salesType;
    const { isProcessed, statusMsg } = await robinson.consolidate();
    if (isProcessed) {
        robinson.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

const consolidateMetro = async(store, action, cutOff) => {
    const metro = new Metro();
    metro.chain = store;
    metro.action = action;
    metro.cutOff = cutOff;
    const { isProcessed, statusMsg } = await metro.consolidate();
    if (isProcessed) {
        metro.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

const consolidatePuregold = async(store, action, cutOff) => {
    const puregold = new Puregold();
    puregold.chain = store;
    puregold.action = action;
    puregold.cutOff = cutOff;
    const { isProcessed, statusMsg } = await puregold.consolidate();
    if (isProcessed) {
        puregold.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

const consolidateWeShop = async(store, action, cutOff) => {
    const weshop = new WeShop();
    weshop.chain = store;
    weshop.action = action;
    weshop.cutOff = cutOff;
    const { isProcessed, statusMsg } = await weshop.consolidate();
    if (isProcessed) {
        weshop.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

const consolidateMerrymart = async(store, action, cutOff) => {
    const merrymart = new MerryMart();
    merrymart.chain = store;
    merrymart.action = action;
    merrymart.cutOff = cutOff;
    const { isProcessed, statusMsg } = await merrymart.consolidate();
    if (isProcessed) {
        merrymart.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

const consolidateWaltermart = async(store, action, cutOff) => {
    const waltermart = new WalterMart();
    waltermart.chain = store;
    waltermart.action = action;
    waltermart.cutOff = cutOff;
    const { isProcessed, statusMsg } = await waltermart.consolidate();
    if (isProcessed) {
        waltermart.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

module.exports = { 
    consolidateMetro, 
    consolidateRobinson, 
    consolidatePuregold,
    consolidateWeShop,
    consolidateMerrymart,
    consolidateWaltermart
}