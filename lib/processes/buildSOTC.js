const dotenv = require("dotenv");
dotenv.config();

const { Porkmeat } = require('../classes/porkmeat');
const { Poultry } = require('../classes/poultry');
const { Swine } = require('../classes/swine');

const buildPorkmeat = async(meat, action) => {
    const porkmeat = new Porkmeat();
    porkmeat.meat = meat;
    porkmeat.action = action;
    await porkmeat.clearSOTCPickupDataSheet();
    const { isProcessed, statusMsg } = await porkmeat.buildSOTC();
    if (isProcessed) {
        porkmeat.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

const buildPoultry = async(meat, action) => {
    const poultry = new Poultry();
    poultry.meat = meat;
    poultry.action = action;
    await poultry.clearSOTCPickupDataSheet();
    const { isProcessed, statusMsg } = await poultry.buildSOTC();
    if (isProcessed) {
        poultry.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

const buildSwine = async(meat, action) => {
    const swine = new Swine();
    swine.meat = meat;
    swine.action = action;
    const isCleared = await swine.clearSOTCPickupDataSheet();
    if (isCleared) {
        const { isProcessed, statusMsg } = await swine.buildSOTC();
        if (isProcessed) {
            swine.log();
            console.log(statusMsg);
        } else {
            console.log(statusMsg);
        }
    }
}

module.exports = { buildPorkmeat, buildPoultry, buildSwine }