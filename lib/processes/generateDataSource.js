const dotenv = require("dotenv");
dotenv.config();

const { Porkmeat } = require('../classes/porkmeat');
const { Poultry } = require('../classes/poultry');
const { Swine } = require('../classes/swine');

const generatePorkmeat = async(meat, action) => {    
    const porkmeat = new Porkmeat();
    porkmeat.meat = meat;
    porkmeat.action = action;
    const { isProcessed, statusMsg } = await porkmeat.generateOutputData();
    if (isProcessed) {
        porkmeat.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

const generatePoultry = async(meat, action) => {    
    const poultry = new Poultry();
    poultry.meat = meat;
    poultry.action = action;
    const { isProcessed, statusMsg } = await poultry.generateOutputData();
    if (isProcessed) {
        poultry.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

const generateSwine = async(meat, action) => {    
    const swine = new Swine();
    swine.meat = meat;
    swine.action = action;
    const { isProcessed, statusMsg } = await swine.generateOutputData();
    if (isProcessed) {
        swine.log();
        console.log(statusMsg);
    } else {
        console.log(statusMsg);
    }
}

module.exports = {
    generatePorkmeat,
    generatePoultry,
    generateSwine,
};