const dotenv = require("dotenv");
dotenv.config();

const ExcelJS = require('exceljs');
const fs = require('fs');
const { Log } = require('./logs');
const { DataFiles } = require('./files');
const { appLabels } = require('../contants/contants');
const { startsWithZero, removeLeadingZero } = require('../utils/utils');

class WalterMart {
    constructor() {
        this.chain = null;
        this.action = null;
        this.cutOff = null;
    }

    setChain(chain) { this.chain = chain; }
    getChain() { return this.chain; }

    setAction(action) { this.action = action; }
    getAction() { return this.action; }

    setAction(cutOff) { this.cutOff = cutOff; }
    getAction() { return this.cutOff; }    

    log() {
        const log = new Log();
        log.filePath = `${process.env.LOG_FILE}`;
        log.chain = this.chain;
        log.action = this.action;
        log.logActivity();
    }

    async processGeneration(filename) {
        try {
            const currentDate = new Date();

            const sourceFile = `${process.env.RAW_DATA_WALTERMART}/${filename}`;
            const sourceSheetName = 1;
            const sourceWB = new ExcelJS.Workbook();

            return await sourceWB.xlsx.readFile(sourceFile).then(() => {
                const sourceSheet = sourceWB.getWorksheet(sourceSheetName);
                
                const destinationWB = new ExcelJS.Workbook();
                destinationWB.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(async() => {
                    const destinationSheet = destinationWB.getWorksheet(`${process.env.CON_SHEET_WALTERMART}`);

                    const showcaseSheet = destinationWB.getWorksheet(`${process.env.STORE_SHEET_SHOWCASE}`);
                    const srpSheet = destinationWB.getWorksheet(`${process.env.STORE_SHEET_SRP}`);
                    const vamSheet = destinationWB.getWorksheet(`${process.env.STORE_SHEET_VAM}`);

                    const consolidatedSheet = destinationWB.getWorksheet(`${process.env.SKU_SHEET_CONSOLIDATED}`);
                    const commrateSheet = destinationWB.getWorksheet(`${process.env.SKU_SHEET_COMMRATE}`);
                    const ninersSheet = destinationWB.getWorksheet(`${process.env.SKU_SHEET_NINERS}`);

                    sourceSheet.eachRow({ includeEmpty: false, firstRow: 3 }, (row, rowNumber) => {
                        const rowData = [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18].map(col => row.getCell(col).value);
                        if (rowNumber > 3) {
                            const cutOffSegments = this.cutOff.split(' ');
                            const cutOffValue = `${cutOffSegments[0]} ${cutOffSegments[1]} to ${cutOffSegments[3]}`.toUpperCase(); 

                            const newRowData = [
                                currentDate.getFullYear(), // YEAR
                                cutOffSegments[0].toUpperCase(), // MONTH
                                cutOffValue, // CUT-OFF
                                this.chain, // CHAIN
                                this.chain, // BANNER
                                parseInt(rowData[0].result), // VENDOR CODE
                                parseInt(rowData[3]), // STORE CODE
                                rowData[4], // STORE NAME
                                parseInt(rowData[5]), // SKU#
                                parseInt(rowData[6].result.trim()), // BARCODE
                                rowData[7], // DESCRIPTION
                                rowData[8], // SELLING UOM
                                parseFloat(rowData[9]).toFixed(5), // ORIG QTY
                                "-", // PACK
                                "-", // KILOS
                                parseFloat(0).toFixed(5), // PCS
                                parseFloat(rowData[10]).toFixed(5), // SALES AMOUNT
                                parseFloat(rowData[11]).toFixed(5), // COMMISSION
                                parseFloat(rowData[12]).toFixed(5), // PAYABLE AMOUNT
                                "-", // NOTE
                                "-", // SKU CATEGORY
                                "-", // AREA
                                "-", // KAM
                                "-", // SKU PER BRAND
                                "-", // GENERALIZED SKU
                                "-", // MOTHER SKU
                                "RETAIL", // SALES CATEGORY
                                "-", // SKU DEPARTMENT
                                "-", // PLACEMENT
                                "-", // PLACEMENT MARKS
                            ];
                            destinationSheet.addRow(newRowData);
                        }
                    });
                    await destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);

                    destinationSheet.eachRow({ includeEmpty: false, firstRow: 2}, (row, rowNumber) => {
                        if (rowNumber > 1) {
                            row.getCell(6).alignment = { horizontal: 'right' };  // VENDOR CODE
                            row.getCell(7).alignment = { horizontal: 'right' };  // STORE CODE
                            row.getCell(9).alignment = { horizontal: 'right' };  // SKU#
                            row.getCell(10).alignment = { horizontal: 'right' };  // BARCODE
                            row.getCell(13).numFmt = `###0.00000`;  // ORIG QTY
                            row.getCell(13).alignment = { horizontal: 'right' };                             
                            row.getCell(14).numFmt = `###0.00000`;  // PACK
                            row.getCell(14).alignment = { horizontal: 'right' };
                            row.getCell(14).value = { formula: `IF(L${rowNumber}="PCK", M${rowNumber}, 0)`};    
                            row.getCell(15).numFmt = `###0.00000`; // KILOS
                            row.getCell(15).alignment = { horizontal: 'right' };
                            row.getCell(15).value = { formula: `IF(L${rowNumber}="KGS", M${rowNumber}, IF(L${rowNumber}="PCK", M${rowNumber}*VLOOKUP(I${rowNumber},Sku_Consolidated!A2:U${consolidatedSheet.lastRow.number},21, FALSE),M${rowNumber}*VLOOKUP(I${rowNumber},Sku_Consolidated!A2:U${consolidatedSheet.lastRow.number},21, FALSE)))`}
                            row.getCell(16).numFmt = `###0.00000`; // PCS
                            row.getCell(16).alignment = { horizontal: 'right' }; 
                            row.getCell(17).alignment = { horizontal: 'right' };  // SALES AMOUNT
                            row.getCell(18).alignment = { horizontal: 'right' };  // COMMISSION
                            row.getCell(19).alignment = { horizontal: 'right' };  // PAYABLE AMOUNT
                            // NOTE
                            row.getCell(21).value = { formula: `VLOOKUP(I${rowNumber},Sku_Consolidated!A2:G${consolidatedSheet.lastRow.number},7, FALSE)`}; // SKU CATEGORY
                            row.getCell(22).value = { formula: `IF(IFERROR(VLOOKUP(G${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},7, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(G${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},7, FALSE), TRUE)=TRUE,VLOOKUP(G${rowNumber},Store_VAM!B2:H${vamSheet.lastRow.number},7, FALSE),VLOOKUP(G${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},7, FALSE)), VLOOKUP(G${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},7, FALSE))`}; // AREA
                            row.getCell(23).value = { formula: `IF(IFERROR(VLOOKUP(G${rowNumber},Store_Showcase!B2:I${showcaseSheet.lastRow.number},8, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(G${rowNumber},Store_SRP!B2:I${srpSheet.lastRow.number},8, FALSE), TRUE)=TRUE,VLOOKUP(G${rowNumber},Store_VAM!B2:I${vamSheet.lastRow.number},8, FALSE),VLOOKUP(G${rowNumber},Store_SRP!B2:I${srpSheet.lastRow.number},8, FALSE)), VLOOKUP(G${rowNumber},Store_Showcase!B2:I${showcaseSheet.lastRow.number},8, FALSE))`}; // KAM
                            row.getCell(24).value = { formula: `VLOOKUP(I${rowNumber},Sku_Consolidated!A2:W${consolidatedSheet.lastRow.number},11,FALSE)`}; // SKU PER BRAND
                            row.getCell(25).value = { formula: `VLOOKUP(I${rowNumber},Sku_Consolidated!A2:H${consolidatedSheet.lastRow.number},8,FALSE)`}; // GENERALIZED SKU  
                            row.getCell(26).value = { formula: `VLOOKUP(I${rowNumber},Sku_Consolidated!A2:I${consolidatedSheet.lastRow.number},9,FALSE)`}; // MOTHER SKU
                            // SALES CATEGORY
                            row.getCell(28).value = { formula: `VLOOKUP(I${rowNumber},Sku_Consolidated!A2:E${consolidatedSheet.lastRow.number},5,FALSE)`}; // SKU DEPT  
                            row.getCell(29).value = { formula: `IF(VLOOKUP(I${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)="SHOWCASE", VLOOKUP(G${rowNumber},Store_Showcase!B2:N${showcaseSheet.lastRow.number},13, FALSE), IF(VLOOKUP(I${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)="VAM",VLOOKUP(G${rowNumber},Store_VAM!B2:N${vamSheet.lastRow.number},13, FALSE),VLOOKUP(G${rowNumber},Store_SRP!B2:N${srpSheet.lastRow.number},13, FALSE)))`}; // PLACEMENT
                            row.getCell(30).value = { formula: `IF(IFERROR(AC${rowNumber},TRUE)=TRUE, "-","OK")`}; // PLACEMENT REMARKS
                            row.getCell(31).value = { formula: `IF(IFERROR(VLOOKUP(G${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},5, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(G${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},5, FALSE), TRUE)=TRUE,VLOOKUP(G${rowNumber},Store_VAM!B2:H${vamSheet.lastRow.number},5, FALSE),VLOOKUP(G${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},5, FALSE)), VLOOKUP(G${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},5, FALSE))`}; // CITY
                            row.getCell(32).value = { formula: `IF(IFERROR(VLOOKUP(G${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},6, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(G${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},6, FALSE), TRUE)=TRUE,VLOOKUP(G${rowNumber},Store_VAM!B2:H${vamSheet.lastRow.number},6, FALSE),VLOOKUP(G${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},6, FALSE)), VLOOKUP(G${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},6, FALSE))`}; // PROVINCE
                            row.getCell(33).value = { formula: `VLOOKUP(I${rowNumber},Sku_Consolidated!A2:R${consolidatedSheet.lastRow.number},18,FALSE)`}; // SKU IDENTIFIER 1
                            row.getCell(34).value = { formula: `VLOOKUP(I${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)`}; // SKU IDENTIFIER 2
                            row.getCell(35).value = { formula: `VLOOKUP(I${rowNumber},Sku_Consolidated!A2:T${consolidatedSheet.lastRow.number},20,FALSE)`}; // INTERNAL BRAND
                        }
                    });

                    destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`).then(() => {
                        const fileManager = new DataFiles();
                        fileManager.copyFile(`${process.env.OUTPUT_FILE}`,`${process.env.OUTPUT_FILE_WALTERMART}`);

                        this.checkFileExists((err, exists) => {
                            if (err) {
                                console.error('Error:', err.message);
                            } else {
                                this.clearOutputDataSheet(destinationWB);
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

    clearOutputDataSheet(workbook) {
        workbook.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(() => {
            const clearsheet = workbook.getWorksheet(`${process.env.CON_SHEET_WALTERMART}`);
            const rowCount = clearsheet.rowCount;
            for (let i = rowCount; i > 1; i--) { clearsheet.spliceRows(i, 1); }                                
            workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);  
            
            this.removeUnrelatedSheets();
        });
    }    

    removeUnrelatedSheets() {
        const workbook = new ExcelJS.Workbook();
        workbook.xlsx.readFile(`${process.env.OUTPUT_FILE_WALTERMART}`).then(() => {
            workbook.eachSheet(sheet => {
                if (!sheet.name.startsWith('Sku_') && !sheet.name.startsWith('Store_') && sheet.name !== `${process.env.CON_SHEET_WALTERMART}`) {
                    workbook.removeWorksheet(sheet.id);
                }                        
            });
            return workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE_WALTERMART}`);
        })
    }     
    
    buildRawData() {
        try {
            return true;

        } catch(e) {
            console.log(e)
            return false;
        }
    }

    async generateOutputData() {
        try {
            const chain = this.chain;
            const fileManager = new DataFiles();
            fileManager.source = process.env.RAW_DATA_WALTERMART;
            const files = fileManager.listFiles().filter(f => f.includes('xlsx'));
            if (files.length > 0) {
                let processResult = [];
                const promises = files.map(async(file) => {
                    return await this.processGeneration(file).then((item) => {
                        let isCompleted = item;
                        if (isCompleted) {
                            fileManager.destination = `${process.env.RAW_DATA_WALTERMART}/${process.env.PROCESSED}`;
                            fileManager.filename = file.trim();
                            fileManager.moveFile();
                        }
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
                            statusMsg: `${chain} - ${appLabels.chainMsg}`
                        }
                    }
                });
                
            } else {
                return {
                    isProcessed: false,
                    statusMsg: `NO DATA FILE(S) FOUND FROM ${chain}!`
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
        try {
            const sourceFile = `${process.env.OUTPUT_FILE_WALTERMART}`.replace(`${process.env.OUTPUT_DIR}`, `${process.env.TEMPO_DATA_DIR}`).replace('.xlsx', '.csv');
            const sourceWB = new ExcelJS.Workbook();

            const fileManager = new DataFiles();
            fileManager.source = `${process.env.TEMPO_DATA_DIR}`;
            const fileResult = fileManager.listFiles().filter(f => f.includes(`${process.env.CON_SHEET_WALTERMART}`.toLowerCase()));

            if (fileResult.length > 0) {
                return await sourceWB.csv.readFile(sourceFile).then(() => {
                    const sourceSheet = sourceWB.worksheets[0];

                    const destinationWB = new ExcelJS.Workbook();
                    destinationWB.xlsx.readFile(`${process.env.CONSOLIDATED_DATA_FILE}`).then(async() => {
                        const destinationSheet = destinationWB.getWorksheet(`${process.env.CONSOLIDATED_SHEET}`);

                        sourceSheet.eachRow({ includeEmpty: false, firstRow: 2 }, (row, rowNumber) => {
                            if (rowNumber > 1 && rowNumber !== undefined) {
                                const newRowData = [
                                    row.getCell(1).value, // YEAR
                                    row.getCell(2).value, // MONTH
                                    row.getCell(4).value, // CHAIN
                                    row.getCell(5).value, // BANNER
                                    row.getCell(8).value.trim(), // BRANCH
                                    row.getCell(21).value, // SKU CATEGORY
                                    row.getCell(11).value, // DESCRIPTION
                                    parseFloat(row.getCell(14).value).toFixed(5), // PACK
                                    parseFloat(row.getCell(15).value).toFixed(5), // KG
                                    parseFloat(row.getCell(16).value).toFixed(5), // PCS
                                    parseFloat(row.getCell(17).value).toFixed(5), // GROSS
                                    parseFloat(row.getCell(19).value).toFixed(5), // NET SALES
                                    row.getCell(22).value, // AREA
                                    row.getCell(9).value, // SKU NUMBER
                                    row.getCell(24).value, // SKU PER BRAND
                                    row.getCell(25).value, // GENERALIZED SKU
                                    row.getCell(26).value, // MOTHER SKU
                                    row.getCell(27).value, // SALES CATEGORY
                                    row.getCell(28).value, // SKU DEPT
                                    row.getCell(29).value, // PLACEMENT
                                    row.getCell(30).value, // PLACEMENT REMARKS
                                    row.getCell(31).value, // CITY
                                    row.getCell(32).value, // PROVINCE
                                    row.getCell(33).value, // SKU REPORT IDENTIFIER 1
                                    row.getCell(34).value, // SKU REPORT IDENTIFIER 2
                                    "-", // SUKI CO STORE
                                    row.getCell(35).value, // INTERNAL BRAND
                                ];
                                destinationSheet.addRow(newRowData);
                            }
                        });
                        await destinationWB.xlsx.writeFile(`${process.env.CONSOLIDATED_DATA_FILE}`);

                        destinationSheet.eachRow({ includeEmpty: false, firstRow: 2}, (row, rowNumber) => {
                            if (rowNumber > 1) {
                                row.getCell(7).alignment = { horizontal: 'right' };
                                row.getCell(8).alignment = { horizontal: 'right' };
                                row.getCell(9).alignment = { horizontal: 'right' };
                                row.getCell(10).alignment = { horizontal: 'right' };
                                row.getCell(11).alignment = { horizontal: 'right' };
                            }
                        });
                        await destinationWB.xlsx.writeFile(`${process.env.CONSOLIDATED_DATA_FILE}`);
                    });

                }).then(async() => {
                    return {
                        isProcessed: true,
                        statusMsg: `${this.chain}: ${appLabels.consolidationMsg}`
                    }
                }).catch(async(err) => {
                    return {
                        isProcessed: false,
                        statusMsg: `${err}`
                    } 
                });
                
            } else {
                return {
                    isProcessed: false,
                    statusMsg: `NO OUTPUT DATA FILE TO CONSOLIDATE FROM ${this.chain}!`
                } 
            }

        } catch(e) {
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
            fs.access(`${process.env.OUTPUT_FILE_WALTERMART}`, fs.constants.F_OK, (err) => {
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

module.exports = { WalterMart }