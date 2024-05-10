const dotenv = require("dotenv");
dotenv.config();

const ExcelJS = require('exceljs');
const fs = require('fs')
const { Log } = require('./logs');
const { DataFiles } = require('./files');
const { appLabels } = require('../contants/contants');
const { startsWithZero, removeLeadingZero } = require('../utils/utils');

class Porkmeat {
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
                    const destinationSheet = destinationWB.getWorksheet(`${process.env.CON_SHEET_PORKMEAT}`);

                    sourceSheet.eachRow({ includeEmpty: false, firstRow: 2 }, (row, rowNumber) => {
                        if (rowNumber >  1) {
                            if (!row.getCell(14).value.includes("14") && row.getCell(28).value.toLowerCase() === this.meat.toLowerCase()) {                            
                                const journalEntryDate = new Date(row.getCell(15).value);
                                const dateOptions = {weekday: 'long', year: 'numeric', month: 'long', day: 'numeric',};                            
                                const month = journalEntryDate.toLocaleDateString(undefined, dateOptions).split(" ")[1].trim().toUpperCase();
                                
                                let dateValue = journalEntryDate.toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: '2-digit'}).split(" ");
                                dateValue = `${dateValue[1].slice(0, -1)}-${dateValue[0]}-${dateValue[2]}`;
    
                                let salesAmount = (row.getCell(9).value < 0) ? Math.abs(row.getCell(9).value) : row.getCell(9).value * -1;
    
                                const newRowData = [
                                    journalEntryDate.getFullYear(), // YEAR
                                    month, // MONTH
                                    dateValue, // DATE
                                    row.getCell(20).value, // INV NO
                                    row.getCell(21).value, // SO
                                    row.getCell(12).value, // COMPLETE CUSTOMER NAME
                                    "-", // INVTY
                                    row.getCell(16).value, // ITEM
                                    row.getCell(17).value, // ITEM DESCRIPTION
                                    row.getCell(17).value, // MOTHER SKU
                                    row.getCell(24).value.toFixed(3), // QTY
                                    row.getCell(25).value, // UOM
                                    "-", // BOX
                                    row.getCell(24).value.toFixed(3), // KG
                                    "-", // PCS
                                    "-", // PKS
                                    "-", // SET
                                    salesAmount.toFixed(3), // SALES AMOUNT
                                    "-", // HEAD
                                    "-", // PRIMAL
                                    "-", // KAM
                                    "-", // UPDATED CHANNEL
                                    "-", // PRODUCT CATEGORY
                                    "-", // ACCOUNTING CHANNEL
                                ]
                                console.log(newRowData);
                                destinationSheet.addRow(newRowData);
                            }
                        }
                    });
                    await destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);
                });

            }).then(async() => {
                return await true;
            }).catch(async(err) => {
                console.error(err);
                return await false;                
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
                        return item;

                    }).then((res) => {
                        processResult.push(res);
                        return res;

                    }).catch((error) => {
                        console.log(error)
                        return false;
                    });
                });
                return Promise.all(promises).then(function(results) {
                    // console.log(results)
                    if (results.includes(true)) {
                        return {
                            isProcessed: true,
                            statusMsg: `${meat}:  - ${appLabels.dataSourceMsg}`
                        }
                    }
                });                

            } else {
                return {
                    isProcessed: false,
                    statusMsg: `NO SAP EXPORT FILE FOUND!`
                }
            }
        } catch(e) {
            return {
                isProcessed: false,
                statusMsg: e
            }
        }
    }

    async consolidate() {

    }
}

module.exports = { Porkmeat }