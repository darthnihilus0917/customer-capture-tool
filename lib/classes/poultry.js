const dotenv = require("dotenv");
dotenv.config();

const ExcelJS = require('exceljs');
const fs = require('fs')
const { Log } = require('./logs');
const { DataFiles } = require('./files');
const { appLabels } = require('../contants/contants');
const { startsWithZero, removeLeadingZero } = require('../utils/utils');

class Poultry {
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

    async processGeneration(filename) {
        try {
            const sourceFile = `${process.env.RAW_DATA_SAP}/${filename}`;
            const sourceWB = new ExcelJS.Workbook();

            return await sourceWB.xlsx.readFile(sourceFile).then(() => {
                const sourceSheet = sourceWB.worksheets[1];

                const destinationWB = new ExcelJS.Workbook();
                destinationWB.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(async() => {
                    const destinationSheet = destinationWB.getWorksheet(`${process.env.CON_SHEET_POULTRY}`);
                });
            });

        } catch(err) {
            console.error(err);
            return false;
        }
    }

    async generateOutputData() {
        try {
            const meat = this.meat;
            const fileManager = new DataFiles();
            fileManager.source = process.env.RAW_DATA_SAP;
            const files = fileManager.listFiles().filter(f => f.includes('.xlsx') && !f.includes('~'));
            if (files.length > 0) {
                let processResult = [];
                const promises = files.map(async(file) => {
                    return await this.processGeneration(file).then((item) => {
                        return true;

                    }).then((res) => {
                        processResult.push(res);
                        return res;

                    }).catch((error) => {
                        console.log(error)
                        return false;
                    });
                });
                return Promise.all(promises).then(function(results) {
                    if (results.includes(true)) {
                        return {
                            isProcessed: true,
                            statusMsg: `${meat}: ${appLabels.dataSourceMsg}`
                        }
                    }
                });                

            } else {
                return {
                    isProcessed: false,
                    statusMsg: `${appLabels.noSapFile.toUpperCase()}`
                }
            }            

        } catch(e) {
            return {
                isProcessed: false,
                statusMsg: e
            }
        }
    }

    async consolidate() {}

    async buildSOTC() {
        try {
            const meat = this.meat;
            const fileManager = new DataFiles();
            fileManager.source = process.env.SOTC_FILE_POULTRY;
            const files = fileManager.listFiles().filter(f => f.includes('.xlsx') && !f.includes('~'));

            if (files.length > 1) {
                return {
                    isProcessed: false,
                    statusMsg: `${appLabels.tooManyFiles}`
                }
            }

            const sourceFile = `${process.env.SOTC_FILE_POULTRY}/${files[0]}`;
            const sourceSOTCSheet = `${process.env.SOTC_SHEET}`;
            const sourcePickupSheet = `${process.env.PICKUP_SHEET}`;
            const sourceWB = new ExcelJS.Workbook();

            // SOTC & PICKUP DATA BUILDUP
            return await sourceWB.xlsx.readFile(sourceFile).then(() => {
                const SOTCSheet = sourceWB.getWorksheet(sourceSOTCSheet);
                const pickupSheet = sourceWB.getWorksheet(sourcePickupSheet);

                const destinationWB = new ExcelJS.Workbook();
                destinationWB.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(async() => {
                    const destinationSOTCSheet = destinationWB.getWorksheet(`${process.env.SOTC_SHEET_POULTRY}`);
                    const destinationPickupSheet = destinationWB.getWorksheet(`${process.env.PICKUP_SHEET_POULTRY}`);

                    SOTCSheet.eachRow({ includeEmpty: false, firstRow: 4 }, (row, rowNumber) => {
                        if (rowNumber > 4) {
                            if (row.getCell(4).value !== null && row.getCell(4).value !== 'STO') {
                            
                                const newRowData = [
                                    parseInt(row.getCell(4).value),
                                    row.getCell(14).value,
                                ]
                                destinationSOTCSheet.addRow(newRowData);
                            }
                        }
                    });

                    pickupSheet.eachRow({ includeEmpty: false, firstRow: 4 }, (row, rowNumber) => {
                        if (rowNumber > 3) {
                            if (row.getCell(4).value !== null && row.getCell(4).value !== 'STO') {
                            
                                const newRowData = [
                                    row.getCell(4).value,
                                    row.getCell(14).value,
                                ]
                                destinationPickupSheet.addRow(newRowData);
                            }
                        }
                    });

                    destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`).then(() => {
                        return true;

                    }).then(() => {
                        return true;

                    }).catch((err) => {
                        console.error(err);
                        return false;
                    });
                });

            }).then(async() => {
                return {
                    isProcessed: true,
                    statusMsg: `${meat}: ${appLabels.sotcDataMsg}`
                }

            }).catch(async(err) => {
                return {
                    isProcessed: false,
                    statusMsg: err
                }
            });

        } catch (e) {
            return {
                isProcessed: false,
                statusMsg: e
            }
        }
    }

    checkFileExists(callback) {
        let attempts = 0;
        const maxAttempts = 3;
        const delay = 1000; // Delay in milliseconds between each attempt
    
        function check() {
            fs.access(`${process.env.OUTPUT_FILE_POULTRY}`, fs.constants.F_OK, (err) => {
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

module.exports = { Poultry }