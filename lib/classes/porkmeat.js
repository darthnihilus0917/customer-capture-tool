const dotenv = require("dotenv");
dotenv.config();

const ExcelJS = require('exceljs');
const fs = require('fs')
const { Log } = require('./logs');
const { DataFiles } = require('./files');
const { appLabels } = require('../contants/contants');
// const { startsWithZero, removeLeadingZero } = require('../utils/utils');

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

                    const sotcSheet = destinationWB.getWorksheet(`${process.env.SOTC_SHEET_PORKMEAT}`);
                    const pickupSheet = destinationWB.getWorksheet(`${process.env.PICKUP_SHEET_PORKMEAT}`);
                    const customerSheet = destinationWB.getWorksheet(`${process.env.CUSTOMER_SHEET_PORKMEAT}`);
                    const skuSheet = destinationWB.getWorksheet(`${process.env.SKU_SHEET_PORKMEAT}`);

                    sourceSheet.eachRow({ includeEmpty: false, firstRow: 2 }, (row, rowNumber) => {
                        if (rowNumber >  1) {
                            if (!row.getCell(14).value.includes("14") && !row.getCell(12).value.includes("POS") 
                                && row.getCell(28).value.toLowerCase() === this.meat.toLowerCase()) {

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
                                    parseInt(row.getCell(21).value), // SO
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
                                    row.getCell(14).value, // ACCOUNTING CHANNEL
                                    row.getCell(12).value, // COMPLETE CUSTOMER NAME
                                ];
                                destinationSheet.addRow(newRowData);
                            }
                        }
                    });
                    await destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);

                    destinationSheet.eachRow({ includeEmpty: false, firstRow: 2}, (row, rowNumber) => {
                        if (rowNumber > 1) {
                            row.getCell(5).alignment = { horizontal: 'left' }; // SO

                            // COMPLETE CUSTOMER NAME
                            const finalCustomerValue = `UPPER(IF(IFERROR(VLOOKUP(E${rowNumber},SOTC_PORKMEAT!A2:B${sotcSheet.lastRow.number},{2},FALSE), TRUE)=TRUE, VLOOKUP(E${rowNumber},PICKUP_PORKMEAT!A2:B${pickupSheet.lastRow.number},{2},FALSE),VLOOKUP(E${rowNumber},SOTC_PORKMEAT!A2:B${sotcSheet.lastRow.number},{2},FALSE)))`;
                            if (row.getCell(6).value === 'ONE TIME CUSTOMER') {
                                row.getCell(6).value = { formula: `VLOOKUP(${finalCustomerValue}, CUSTOMERS_PORKMEAT!A2:B${customerSheet.lastRow.number},{2},FALSE)`}
                            }
                            
                            row.getCell(11).alignment = { horizontal: 'right' }; // QTY
                            row.getCell(14).alignment = { horizontal: 'right' }; // KG
                            row.getCell(18).alignment = { horizontal: 'right' }; // SALES AMOUNT
                             
                            // HEAD
                            row.getCell(19).alignment = { horizontal: 'right' };
                            row.getCell(19).numFmt = `#,##0.000`;
                            row.getCell(19).value = { formula: `R${rowNumber}/N${rowNumber} * -1`}
                            
                            row.getCell(20).value = { formula: `VLOOKUP(H${rowNumber},SKU_PORKMEAT!A2:D${skuSheet.lastRow.number},{3},FALSE)`} // PRIMAL
                            row.getCell(21).value = { formula: `VLOOKUP(${finalCustomerValue}, CUSTOMERS_PORKMEAT!A2:C${customerSheet.lastRow.number},{3},FALSE)`} // KAM
                            row.getCell(22).value = { formula: `VLOOKUP(${finalCustomerValue}, CUSTOMERS_PORKMEAT!A2:D${customerSheet.lastRow.number},{4},FALSE)`} // UPDATED CHANNEL
                            row.getCell(23).value = { formula: `VLOOKUP(H${rowNumber},SKU_PORKMEAT!A2:D${skuSheet.lastRow.number},{4},FALSE)`} // PRODUCT CATEGORY
                        }
                    });

                    destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`).then(() => {
                        const fileManager = new DataFiles();
                        fileManager.copyFile(`${process.env.OUTPUT_FILE}`,`${process.env.OUTPUT_FILE_PORKMEAT}`);
                        this.checkFileExists(process.env.OUTPUT_FILE_PORKMEAT, (err, exists) => {
                            if (err) {
                                console.error('Error:', err.message);
                            } else {
                                this.clearOutputDataSheet(process.env.CON_SHEET_PORKMEAT, destinationWB);
                            }
                        });

                    }).then(() => {
                        return true;
                    }).catch((error) => {
                        console.error(error);
                        return false;
                    }); 
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

    clearSOTCPickupDataSheet(workbook) {
        workbook.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(() => {
            const sotcSheet = workbook.getWorksheet(`${process.env.SOTC_SHEET_PORKMEAT}`);
            const pickupSheet = workbook.getWorksheet(`${process.env.PICKUP_SHEET_PORKMEAT}`);

            const sotcCount = sotcSheet.rowCount;
            for (let i = sotcCount; i > 1; i--) { sotcSheet.spliceRows(i, 1); }
            workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);
            
            const pickupCount = pickupSheet.rowCount;
            for (let i = pickupCount; i > 1; i--) { pickupSheet.spliceRows(i, 1); }            
            workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);
        });
    }    

    clearOutputDataSheet(sheetname, workbook) {
        workbook.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(() => {
            const clearsheet = workbook.getWorksheet(`${sheetname}`);
            const rowCount = clearsheet.rowCount;
            for (let i = rowCount; i > 1; i--) { clearsheet.spliceRows(i, 1); }                                
            workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);  
            
            this.removeUnrelatedSheets();
        });
    }

    removeUnrelatedSheets() {
        const workbook = new ExcelJS.Workbook();
        workbook.xlsx.readFile(`${process.env.OUTPUT_FILE_PORKMEAT}`).then(() => {
            workbook.eachSheet(sheet => {
                const sheetname = process.env.CON_SHEET_PORKMEAT;
                const sku = `SKU_${sheetname}`;
                const customers = `CUSTOMERS_${sheetname}`;
                const sotc = `SOTC_${sheetname}`;
                const pickup = `PICKUP_${sheetname}`;

                if (!sheet.name.startsWith(sku) && !sheet.name.startsWith(customers) 
                    && !sheet.name.startsWith(sotc) && !sheet.name.startsWith(pickup) && sheet.name !== `${process.env.CON_SHEET_PORKMEAT}`) {
                    workbook.removeWorksheet(sheet.id);
                }
            });
            return workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE_PORKMEAT}`);
        })
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
            fileManager.source = process.env.SOTC_FILE_PORKMEAT;
            const files = fileManager.listFiles().filter(f => f.includes('.xlsx') && !f.includes('~'));

            if (files.length > 1) {
                return {
                    isProcessed: false,
                    statusMsg: `${appLabels.tooManyFiles}`
                }
            }

            const sourceFile = `${process.env.SOTC_FILE_PORKMEAT}/${files[0]}`;
            const sourceSOTCSheet = `${process.env.SOTC_SHEET}`;
            const sourcePickupSheet = `${process.env.PICKUP_SHEET}`;
            const sourceWB = new ExcelJS.Workbook();

            // SOTC & PICKUP DATA BUILDUP
            return await sourceWB.xlsx.readFile(sourceFile).then(() => {
                const SOTCSheet = sourceWB.getWorksheet(sourceSOTCSheet);
                const pickupSheet = sourceWB.getWorksheet(sourcePickupSheet);

                const destinationWB = new ExcelJS.Workbook();
                this.clearSOTCPickupDataSheet(destinationWB);

                destinationWB.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(async() => {
                    const destinationSOTCSheet = destinationWB.getWorksheet(`${process.env.SOTC_SHEET_PORKMEAT}`);
                    const destinationPickupSheet = destinationWB.getWorksheet(`${process.env.PICKUP_SHEET_PORKMEAT}`);

                    SOTCSheet.eachRow({ includeEmpty: false, firstRow: 4 }, (row, rowNumber) => {
                        if (rowNumber > 4) {
                            if (row.getCell(4).value !== null && row.getCell(4).value !== 'STO') {
                            
                                const newRowData = [
                                    parseInt(row.getCell(4).value),
                                    row.getCell(14).value,
                                ];
                                destinationSOTCSheet.addRow(newRowData);
                            }
                        }
                    });

                    pickupSheet.eachRow({ includeEmpty: false, firstRow: 4 }, (row, rowNumber) => {
                        if (rowNumber > 3) {
                            if (row.getCell(4).value !== null && row.getCell(4).value !== 'STO') {
                            
                                const newRowData = [
                                    parseInt(row.getCell(4).value),
                                    row.getCell(14).value,
                                ];
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

    checkFileExists(filePath, callback) {
        let attempts = 0;
        const maxAttempts = 3;
        const delay = 1000; // Delay in milliseconds between each attempt
    
        function check() {
            fs.access(`${filePath}`, fs.constants.F_OK, (err) => {
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

module.exports = { Porkmeat }