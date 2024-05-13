const dotenv = require("dotenv");
dotenv.config();

const ExcelJS = require('exceljs');
const fs = require('fs')
const { Log } = require('./logs');
const { DataFiles } = require('./files');
const { appLabels } = require('../contants/contants');
const { startsWithZero, removeLeadingZero } = require('../utils/utils');

class Swine {
    constructor() {
        this.meat = null;
        this.action = null;
    }

    setChain(meat) { this.meat = meat; }
    getChain() { return this.meat; }

    setAction(action) { this.action = action; }
    getAction() { return this.action; }

    log() {
        const log = new Log();
        log.filePath = `${process.env.LOG_FILE}`;
        log.meat = this.meat;
        log.action = this.action;
        log.logActivity();
    }

    async generateOutputData() {

    }

    async consolidate() {

    }

    async buildSOTC() {
        
    }

    checkFileExists(callback) {
        let attempts = 0;
        const maxAttempts = 3;
        const delay = 1000; // Delay in milliseconds between each attempt
    
        function check() {
            fs.access(`${process.env.OUTPUT_FILE_SWINE}`, fs.constants.F_OK, (err) => {
                if (!err) {
                    // File exists
                    callback(null, true);
                } else {
                    // File does not exist
                    attempts++;
                    if (attempts < maxAttempts) {
                        // Retry after delay
                        setTimeout(check, delay);
                    } else {
                        // Max attempts reached
                        callback(new Error('File does not exist after multiple attempts'), false);
                    }
                }
            });
        }    
        check(); // Start checking
    } 
}

module.exports = { Swine }