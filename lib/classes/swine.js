const dotenv = require("dotenv");
dotenv.config();

const ExcelJS = require('exceljs');
const fs = require('fs')
const { Log } = require('./logs');
const { DataFiles } = require('./files');
const { appLabels } = require('../contants/contants');

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

    async processGeneration(filename) {
        try {
            const sourceFile = `${process.env.RAW_DATA_SAP}/${filename}`;
            const sourceWB = new ExcelJS.Workbook();

            return await sourceWB.xlsx.readFile(sourceFile).then(() => {
                const sourceSheet = sourceWB.worksheets[1];

                const destinationWB = new ExcelJS.Workbook();
                destinationWB.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(async() => {
                    const destinationSheet = destinationWB.getWorksheet(`${process.env.CON_SHEET_SWINE}`);

                    const sotcSheet = destinationWB.getWorksheet(`${process.env.SOTC_SHEET_SWINE}`);
                    const pickupSheet = destinationWB.getWorksheet(`${process.env.PICKUP_SHEET_SWINE}`);
                    const skuSheet = destinationWB.getWorksheet(`${process.env.SKU_SHEET_SWINE}`);

                    sourceSheet.eachRow({ includeEmpty: false, firstRow: 2 }, (row, rowNumber) => {
                        if (rowNumber >  1) {
                            if (!row.getCell(14).value.includes("14") && !row.getCell(12).value.includes("POS") 
                                && row.getCell(28).value.toLowerCase() === 'live') {

                                const journalEntryDate = new Date(row.getCell(15).value);
                                const dateOptions = {weekday: 'long', year: 'numeric', month: 'long', day: 'numeric',};                            
                                const month = journalEntryDate.toLocaleDateString(undefined, dateOptions).split(" ")[1].trim().toUpperCase();

                                let dateValue = journalEntryDate.toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: '2-digit'}).split(" ");
                                dateValue = `${dateValue[1].slice(0, -1)}-${dateValue[0]}-${dateValue[2]}`;

                                let salesAmount = (row.getCell(9).value < 0) ? Math.abs(row.getCell(9).value) : row.getCell(9).value * -1;
                                const soNo = parseInt(row.getCell(21).value);
                                const materialId = parseInt(row.getCell(16).value);
                                const helperRef = `${soNo}${materialId}`;

                                const newRowData = [
                                    journalEntryDate.getFullYear(), // YEAR
                                    month, // MONTH
                                    dateValue, // DATE
                                    row.getCell(20).value, // INV NO
                                    soNo, // SO
                                    row.getCell(12).value, // COMPLETE CUSTOMER NAME
                                    "-", // INVTY
                                    "-", // FARM
                                    materialId, // ITEM
                                    row.getCell(17).value.toUpperCase(), // ITEM DESCRIPTION
                                    row.getCell(17).value.toUpperCase(), // MOTHER SKU
                                    "-", // CLASS
                                    parseFloat(Number(row.getCell(24).value) * -1), // QTY
                                    row.getCell(25).value, // UOM
                                    "-", // GROSS WEIGHT
                                    "-", // DISC
                                    "-", // WEIGHT
                                    "-", // AVERAGE WEIGHT
                                    "-", // VAL
                                    parseFloat(salesAmount), // SALES AMOUNT
                                    "-", // / KILO
                                    "-", // HEAD
                                    row.getCell(12).value.trim(), // COMPLETE CUSTOMER NAME
                                    row.getCell(14).value.trim(), // DIST. CHANNEL
                                    helperRef
                                ];
                                // console.log(newRowData);
                                destinationSheet.addRow(newRowData);
                            }
                        }
                    });
                    await destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);

                    destinationSheet.eachRow({ includeEmpty: false, firstRow: 1}, (row, rowNumber) => {
                        if (rowNumber > 1) {
                            // SO
                            row.getCell(5).alignment = { horizontal: 'left' }; 
                            
                            // COMPLETE CUSTOMER NAME
                            if (row.getCell(6).value === 'ONE TIME CUSTOMER' || row.getCell(6).value === 'WALK-IN') {
                                if (row.getCell(21).value.includes('20')) {
                                    row.getCell(6).value = "FEEDMILL EMPLOYEES";
                                } else {
                                    const customerValue = `IF(ISBLANK(VLOOKUP(E${rowNumber},SOTC_SWINE!B2:D${sotcSheet.lastRow.number},3,FALSE)),VLOOKUP(E${rowNumber},SOTC_SWINE!B2:E${sotcSheet.lastRow.number},4,FALSE),VLOOKUP(E${sotcSheet.lastRow.number},SOTC_SWINE!B2:D${sotcSheet.lastRow.number},3,FALSE))`;
                                    const rightConcat = `"${process.env.PREFIX_SWINE}"&UPPER(RIGHT(${customerValue},LEN(${customerValue})`;
                                    const searchColon = `${rightConcat}-SEARCH(":",${customerValue})))`;
                                    const searchHypen = `${rightConcat}-SEARCH("-",${customerValue})))`;
                                    const finalValue = `IF(ISNUMBER(SEARCH(":",${customerValue}))=TRUE,${searchColon},${searchHypen})`;
                                    row.getCell(6).value = { formula: `${finalValue}`};
                                }                                
                            }

                            // INV
                            row.getCell(7).alignment = { horizontal: 'center' }; 

                            // FARM
                            row.getCell(8).value = { formula: `VLOOKUP(E${rowNumber},SOTC_SWINE!B2:C${sotcSheet.lastRow.number},{2},FALSE)`};
                            row.getCell(8).alignment = { horizontal: 'center' }; 

                            // ITEM
                            row.getCell(9).alignment = { horizontal: 'left' }; 

                            // CLASS
                            const classFormula = `VLOOKUP(K${rowNumber},SKU_SWINE!C2:F${skuSheet.lastRow.number},{4},FALSE)`;
                            row.getCell(12).value = { formula: `IF(${classFormula}=0, "-",${classFormula}`};
                            
                            // QTY
                            row.getCell(13).alignment = { horizontal: 'right' };
                            row.getCell(13).numFmt = `#,##0.000`; 

                            // UOM
                            row.getCell(14).alignment = { horizontal: 'center' }; 

                            // GROSS WEIGHT
                            const qtyKG = `VLOOKUP(Y${rowNumber},SOTC_SWINE!A2:J${sotcSheet.lastRow.number},9,FALSE)`;
                            const qtyHD = `VLOOKUP(Y${rowNumber},SOTC_SWINE!A2:J${sotcSheet.lastRow.number},8,FALSE)`;
                            row.getCell(15).value = { formula: `((${qtyKG}/${qtyHD})*M${rowNumber})`};
                            row.getCell(15).numFmt = `#,##0.000`;

                            // DISCOUNT
                            row.getCell(16).value = (row.getCell(23).value === 'ONE TIME CUSTOMER' || row.getCell(23).value === 'WALK-IN') ? 0 : { formula: `3*M${rowNumber}`};
                            row.getCell(16).alignment = { horizontal: 'right' };
                            row.getCell(16).numFmt = `#,##0.000`;

                            // NET WEIGHT
                            row.getCell(17).value = { formula: `O${rowNumber}-P${rowNumber}`};
                            row.getCell(17).numFmt = `#,##0.000`;

                            // AVERAGE WEIGHT
                            row.getCell(18).value = { formula: `Q${rowNumber}/M${rowNumber}`};
                            row.getCell(18).numFmt = `#,##0.000`;

                            // VAL
                            row.getCell(19).value = { formula: `U${rowNumber}*P${rowNumber}`}; 
                            row.getCell(19).numFmt = `#,##0.000`;

                            // SALES AMOUNT
                            row.getCell(20).alignment = { horizontal: 'right' }; 
                            row.getCell(20).numFmt = `#,##0.000`;
                            
                            // KILO
                            row.getCell(21).value = { formula: `T${rowNumber}/Q${rowNumber}`}; 
                            row.getCell(21).alignment = { horizontal: 'right' };
                            row.getCell(21).numFmt = `#,##0.000`;
                            
                            // HEAD
                            row.getCell(22).value = { formula: `T${rowNumber}/M${rowNumber}`}; 
                            row.getCell(22).alignment = { horizontal: 'right' };
                            row.getCell(22).numFmt = `#,##0.000`;
                        }
                    });

                    destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`).then(() => {
                        const fileManager = new DataFiles();
                        fileManager.copyFile(`${process.env.OUTPUT_FILE}`,`${process.env.OUTPUT_FILE_SWINE}`);
                        this.checkFileExists(process.env.OUTPUT_FILE_SWINE, (err, exists) => {
                            if (err) {
                                console.error('Error:', err.message);
                            } else {
                                this.clearOutputDataSheet(process.env.CON_SHEET_SWINE, destinationWB);
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

    async clearSOTCPickupDataSheet() {
        const workbook = new ExcelJS.Workbook();
        return await workbook.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(async() => {
            const sotcSheet = workbook.getWorksheet(`${process.env.SOTC_SHEET_SWINE}`);

            const clearWorksheet = (worksheet) => { 
                for (let i = worksheet.actualRowCount; i > 1; i--) {
                    const row = worksheet.getRow(i);
                    if (row.hasValues) {
                        worksheet.spliceRows(i, 1);
                    }
                }
            }

            clearWorksheet(sotcSheet);

            await workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);

        }).then(() => {
            console.log(appLabels.sotcCleanUp);
            return true;
        }).catch((error) => {
            console.error(error)
            return false;
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
        workbook.xlsx.readFile(`${process.env.OUTPUT_FILE_SWINE}`).then(() => {
            workbook.eachSheet(sheet => {
                const sheetname = process.env.CON_SHEET_SWINE;
                const sku = `SKU_${sheetname}`;
                const customers = `CUSTOMERS_${sheetname}`;
                const sotc = `SOTC_${sheetname}`;
                const pickup = `PICKUP_${sheetname}`;

                if (!sheet.name.startsWith(sku) && !sheet.name.startsWith(customers) 
                    && !sheet.name.startsWith(sotc) && !sheet.name.startsWith(pickup) && sheet.name !== `${process.env.CON_SHEET_SWINE}`) {
                    workbook.removeWorksheet(sheet.id);
                }
            });
            return workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE_SWINE}`);
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
            fileManager.source = process.env.SOTC_FILE_SWINE;
            const files = fileManager.listFiles().filter(f => f.includes('.xlsx') && !f.includes('~'));

            if (files.length > 1) {
                return {
                    isProcessed: false,
                    statusMsg: `${appLabels.tooManyFiles}`
                }
            }

            const sourceFile = `${process.env.SOTC_FILE_SWINE}/${files[0]}`;
            const sourceWB = new ExcelJS.Workbook();

            // SOTC & PICKUP DATA BUILDUP
            return await sourceWB.xlsx.readFile(sourceFile).then(() => {
                const sourceSheet = sourceWB.worksheets[0];

                const destinationWB = new ExcelJS.Workbook();
                destinationWB.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(async() => {
                    const destinationSOTCSheet = destinationWB.getWorksheet(`${process.env.SOTC_SHEET_SWINE}`);

                    sourceSheet.eachRow({ includeEmpty: false, firstRow: 3 }, (row, rowNumber) => {
                        if (rowNumber > 3) {
                            const soNumber = parseInt(row.getCell(4).value);
                            const materialId = parseInt(row.getCell(28).value);
                            const helperRef = `${soNumber}${materialId}`;

                            const generalObject = {};
                            if (row.getCell(7).value !== null) {
                                generalObject.farm = row.getCell(7).value;
                                generalObject.customerName = row.getCell(17).value;
                                generalObject.customerAddress = row.getCell(18).value;
                            }

                            const newRowData = [
                                helperRef, // HELPER REF
                                soNumber, // SAP SO
                                generalObject.farm, // FARM
                                generalObject.customerName,  // CUSTOMER NAME
                                generalObject.customerAddress, // CUSTOMER ADDRESS
                                materialId, // MATERIAL ID
                                row.getCell(27).value, // MATERIAL DESCRIPTION
                                row.getCell(29).value, // QTY(HD)
                                row.getCell(30).value, // QTY(KG)
                                row.getCell(31).value, // QTY(KG DISCOUNTED)                                
                            ];
                            
                            if (row.getCell(7).value !== null && !row.getCell(7).value.includes('Donation')) {
                                destinationSOTCSheet.addRow(newRowData);
                            } else if (row.getCell(7).value === null) {
                                destinationSOTCSheet.addRow(newRowData);
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
                    // return false;
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

module.exports = { Swine }