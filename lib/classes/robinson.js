const dotenv = require("dotenv");
dotenv.config();

const ExcelJS = require('exceljs');
const fs = require('fs')
const { Log } = require('./logs');
const { DataFiles } = require('./files');
const { appLabels } = require('../contants/contants');
const { startsWithZero, removeLeadingZero } = require('../utils/utils');

class Robinson {
    constructor() {
        this.salesType = null;
        this.chain = null;
        this.action = null;
        this.cutOff = null;
    }

    setChain(chain) { this.chain = chain; }
    getChain() { return this.chain; }

    setSalesType(salesType) { this.salesType = salesType; }
    getSalesType() { return this.salesType; }

    setAction(action) { this.action = action; }
    getAction() { return this.action; }

    setAction(cutOff) { this.cutOff = cutOff; }
    getAction() { return this.cutOff; }

    log() {
        const log = new Log();
        log.filePath = `${process.env.LOG_FILE}`;
        log.chain = this.chain;
        log.salesType = this.salesType;
        log.action = this.action;
        log.logActivity();
    }

    async processGeneration(filename){
        try {
            const currentDate = new Date();

            const sourceFile = (this.salesType === 'RETAIL') ? `${process.env.RAW_DATA_ROBINSON_RETAIL}/${filename}` : `${process.env.RAW_DATA_ROBINSON_ECOMM}/${filename}`
            const sourceSheetName = (this.salesType === 'RETAIL') ? `${process.env.RETAIL_SHEETNAME}` : `${process.env.ECOMM_SHEETNAME}`;        
            const sourceWB = new ExcelJS.Workbook();

            return await sourceWB.xlsx.readFile(sourceFile).then(() => {
                const sourceSheet = sourceWB.getWorksheet(sourceSheetName);

                const destinationWB = new ExcelJS.Workbook();
                destinationWB.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(async() => {
                    const destinationSheet = destinationWB.getWorksheet(`${process.env.CON_SHEET_ROBINSON}`);

                    const showcaseSheet = destinationWB.getWorksheet(`${process.env.STORE_SHEET_SHOWCASE}`);
                    const srpSheet = destinationWB.getWorksheet(`${process.env.STORE_SHEET_SRP}`);
                    const vamSheet = destinationWB.getWorksheet(`${process.env.STORE_SHEET_VAM}`);
        
                    const consolidatedSheet = destinationWB.getWorksheet(`${process.env.SKU_SHEET_CONSOLIDATED}`);
                    const commrateSheet = destinationWB.getWorksheet(`${process.env.SKU_SHEET_COMMRATE}`);
                    const ninersSheet = destinationWB.getWorksheet(`${process.env.SKU_SHEET_NINERS}`);

                    sourceSheet.eachRow({ includeEmpty: false }, (row) => {
                        const rowData = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(col => row.getCell(col).value);
                            
                        if (startsWithZero(rowData[0])) {
                            const cutOffSegments = this.cutOff.split(' ');
                            const packColValue = (rowData[5] === 'PCK' || rowData[5] === 'PCS') ? parseFloat(rowData[6]).toFixed(5) : parseFloat(0).toFixed(5);
                            // const pcsColValue = (rowData[5] === 'PCS') ? parseFloat(rowData[6]).toFixed(5) : parseFloat(0).toFixed(5);
                            const pcsColValue = parseFloat(0).toFixed(5);
                            const kgColValue = (rowData[5] === 'KLS') ? parseFloat(rowData[6]).toFixed(5) : parseFloat(0).toFixed(5);
                            const cutOffValue = `${cutOffSegments[0]} ${cutOffSegments[1]} to ${cutOffSegments[3]}`.toUpperCase(); 
                            
                            const newRowData = [
                                currentDate.getFullYear(), // YEAR
                                cutOffSegments[0].toUpperCase(), // MONTH
                                cutOffValue, // CUT-OFF
                                this.chain, // CHAIN
                                parseInt(removeLeadingZero(rowData[3])), // BRANCH CODE
                                "-", // BRANCH
                                parseInt(removeLeadingZero(rowData[0])), // SKU
                                rowData[1].trim(), // DESCRIPTION
                                parseFloat(rowData[6]).toFixed(5), // ORIG QTY
                                rowData[5], // UOM
                                packColValue, // PACK
                                kgColValue, // KG
                                pcsColValue, // PCS
                                parseFloat(rowData[9]).toFixed(5), // GROSS SALES
                                parseFloat(0).toFixed(5), // COMM AMOUNT
                                parseFloat(rowData[9]).toFixed(5), // NET SALES
                                "-", // SKU CATEGORY
                                "-", // AREA
                                "-", // KAM
                                "-", // SKU REPORT IDENTIFIER 1
                                parseInt(rowData[2]), // UPC
                                "-", // BANNER
                                "-", // SKU PER BRAND
                                "-", // GENERALIZED SKU
                                "-", // MOTHER SKU
                                this.salesType, // SALES CATEGORY
                                "-", // SKU DEPT.
                                "-", // PLACEMENT
                                "-", // PLACEMENT REMARKS
                                parseFloat(rowData[7]).toFixed(5), // NET TAX
                                parseFloat(rowData[8]).toFixed(5) // TAX
                            ]    
                            destinationSheet.addRow(newRowData);
                        }
                    });
                    await destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);

                    destinationSheet.eachRow({ includeEmpty: false, firstRow: 2}, (row, rowNumber) => {
                        if (rowNumber > 1) {
                            row.getCell(6).value = { formula: `IF(IFERROR(VLOOKUP(E${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},{2}, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(E${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},{2}, FALSE), TRUE)=TRUE,VLOOKUP(E${rowNumber},Store_VAM!B2:H${vamSheet.lastRow.number},{2}, FALSE),VLOOKUP(E${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},{2}, FALSE)), VLOOKUP(E${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},{2}, FALSE))`}; // BRANCH
                            row.getCell(9).numFmt = `###0.00000`; // ORIG QTY
                            row.getCell(9).alignment = { horizontal: 'right' };
                            row.getCell(11).numFmt = `###0.00000`; // PACK
                            row.getCell(11).alignment = { horizontal: 'right' }; 
                            // row.getCell(11).value = { formula: `IF(J${rowNumber}="PCK", I${rowNumber}, 0)`};
                            row.getCell(11).value = { formula: `IF(J${rowNumber}="PCK", I${rowNumber}, IF(J${rowNumber}="PCS", I${rowNumber}, 0))`};
                            row.getCell(12).numFmt = `###0.00000`; // KLS
                            row.getCell(12).alignment = { horizontal: 'right' };
                            row.getCell(12).value = { formula: `IF(J${rowNumber}="KLS", I${rowNumber}, IF(J${rowNumber}="PCK", I${rowNumber}*VLOOKUP(G${rowNumber},Sku_Consolidated!A2:U${consolidatedSheet.lastRow.number},21, FALSE),I${rowNumber}*VLOOKUP(G${rowNumber},Sku_Consolidated!A2:U${consolidatedSheet.lastRow.number},21, FALSE)))`}
                            row.getCell(13).numFmt = `###0.00000`; // PCS
                            row.getCell(13).alignment = { horizontal: 'right' };
                            // row.getCell(13).value = { formula: `IF(J${rowNumber}="PCS", I${rowNumber}, 0)`};
                            row.getCell(14).numFmt = `###0.00000`; // GROSS SALES
                            row.getCell(14).alignment = { horizontal: 'right' };
                            row.getCell(15).numFmt = `###0.00000`; // COMM RATE
                            row.getCell(15).alignment = { horizontal: 'right' };
                            row.getCell(15).value = { formula: `N${rowNumber} * VLOOKUP(E${rowNumber}&VLOOKUP(G${rowNumber},Sku_Consolidated!A2:Q${consolidatedSheet.lastRow.number},17,FALSE),Sku_CommRate!A2:E${commrateSheet.lastRow.number}, 5, FALSE)`}
                            row.getCell(16).numFmt = `###0.00000`; // NET SALES
                            row.getCell(16).alignment = { horizontal: 'right' };
                            row.getCell(16).value = { formula: `N${rowNumber} - O${rowNumber}`};
                            row.getCell(17).value = { formula: `VLOOKUP(G${rowNumber},Sku_Consolidated!A2:G${consolidatedSheet.lastRow.number},7, FALSE)`}; // SKU CATEGORY
                            row.getCell(18).value = { formula: `IF(IFERROR(VLOOKUP(E${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},7, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(E${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},7, FALSE), TRUE)=TRUE,VLOOKUP(E${rowNumber},Store_VAM!B2:H${vamSheet.lastRow.number},7, FALSE),VLOOKUP(E${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},7, FALSE)), VLOOKUP(E${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},7, FALSE))`}; // AREA
                            row.getCell(19).value = { formula: `IF(IFERROR(VLOOKUP(E${rowNumber},Store_Showcase!B2:I${showcaseSheet.lastRow.number},8, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(E${rowNumber},Store_SRP!B2:I${srpSheet.lastRow.number},8, FALSE), TRUE)=TRUE,VLOOKUP(E${rowNumber},Store_VAM!B2:I${vamSheet.lastRow.number},8, FALSE),VLOOKUP(E${rowNumber},Store_SRP!B2:I${srpSheet.lastRow.number},8, FALSE)), VLOOKUP(E${rowNumber},Store_Showcase!B2:I${showcaseSheet.lastRow.number},8, FALSE))`}; // KAM
                            row.getCell(20).value = { formula: `VLOOKUP(G${rowNumber},Sku_Consolidated!A2:R${consolidatedSheet.lastRow.number},18,FALSE)`}; // SKU IDENTIFIER 1
                            row.getCell(22).value = { formula: `IF(VLOOKUP(G${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)="SHOWCASE", VLOOKUP(E${rowNumber},Store_Showcase!B2:L${showcaseSheet.lastRow.number},11, FALSE), IF(VLOOKUP(G${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)="VAM",VLOOKUP(E${rowNumber},Store_VAM!B2:L${vamSheet.lastRow.number},11, FALSE),VLOOKUP(E${rowNumber},Store_SRP!B2:L${srpSheet.lastRow.number},11, FALSE)))`}; // BANNER
                            row.getCell(23).value = { formula: `VLOOKUP(G${rowNumber},Sku_Consolidated!A2:W${consolidatedSheet.lastRow.number},11,FALSE)`}; // SKU PER BRAND
                            row.getCell(24).value = { formula: `VLOOKUP(G${rowNumber},Sku_Consolidated!A2:H${consolidatedSheet.lastRow.number},8,FALSE)`}; // GENERALIZED SKU
                            row.getCell(25).value = { formula: `VLOOKUP(G${rowNumber},Sku_Consolidated!A2:I${consolidatedSheet.lastRow.number},9,FALSE)`}; // MOTHER SKU
                            row.getCell(27).value = { formula: `VLOOKUP(G${rowNumber},Sku_Consolidated!A2:E${consolidatedSheet.lastRow.number},5,FALSE)`}; // SKU DEPT
                            row.getCell(28).value = { formula: `IF(VLOOKUP(G${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)="SHOWCASE", VLOOKUP(E${rowNumber},Store_Showcase!B2:N${showcaseSheet.lastRow.number},13, FALSE), IF(VLOOKUP(G${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)="VAM",VLOOKUP(E${rowNumber},Store_VAM!B2:N${vamSheet.lastRow.number},13, FALSE),VLOOKUP(E${rowNumber},Store_SRP!B2:N${srpSheet.lastRow.number},13, FALSE)))`}; // PLACEMENT
                            row.getCell(29).value = { formula: `IF(IFERROR(AB${rowNumber},TRUE)=TRUE, "-","OK")`}; // PLACEMENT REMARKS
                            row.getCell(30).numFmt = `###0.00000`; // NET TAX
                            row.getCell(30).alignment = { horizontal: 'right' };
                            row.getCell(31).numFmt = `###0.00000`; // TAX
                            row.getCell(31).alignment = { horizontal: 'right' };
                            row.getCell(32).value = { formula: `IF(IFERROR(VLOOKUP(E${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},{5}, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(E${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},{5}, FALSE), TRUE)=TRUE,VLOOKUP(E${rowNumber},Store_VAM!B2:H${vamSheet.lastRow.number},{5}, FALSE),VLOOKUP(E${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},{5}, FALSE)), VLOOKUP(E${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},{5}, FALSE))`}; // CITY
                            row.getCell(33).value = { formula: `IF(IFERROR(VLOOKUP(E${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},{6}, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(E${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},{6}, FALSE), TRUE)=TRUE,VLOOKUP(E${rowNumber},Store_VAM!B2:H${vamSheet.lastRow.number},{6}, FALSE),VLOOKUP(E${rowNumber},Store_SRP!B2:H${srpSheet.lastRow.number},{6}, FALSE)), VLOOKUP(E${rowNumber},Store_Showcase!B2:H${showcaseSheet.lastRow.number},{6}, FALSE))`}; // PROVINCE
                            row.getCell(34).value = { formula: `VLOOKUP(G${rowNumber},Sku_Consolidated!A2:R${consolidatedSheet.lastRow.number},18,FALSE)`}; // SKU IDENTIFIER 1
                            row.getCell(35).value = { formula: `VLOOKUP(G${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)`}; // SKU IDENTIFIER 2
                            row.getCell(36).value = { formula: `VLOOKUP(G${rowNumber},Sku_Consolidated!A2:T${consolidatedSheet.lastRow.number},20,FALSE)`}; // INTERNAL BRAND
                        }
                    });
                    
                    destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`).then(() => {
                        const fileManager = new DataFiles();
                        if (this.salesType === 'RETAIL') {
                            fileManager.copyFile(`${process.env.OUTPUT_FILE}`,`${process.env.OUTPUT_FILE_ROBINSON}`);    
                        } else {
                            fileManager.copyFile(`${process.env.OUTPUT_FILE}`,`${process.env.OUTPUT_FILE_ROBINSON_ECOMM}`);
                        }

                        this.checkFileExists(this.salesType, (err, exists) => {
                            if (err) {
                                console.error('Error:', err.message);
                            } else {
                                this.clearOutputDataSheet(destinationWB);
                            }
                        });

                    }).then(() => {
                        return true;
                    }).catch((err) => {
                        console.error(err);
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
            const clearsheet = workbook.getWorksheet(`${process.env.CON_SHEET_ROBINSON}`);
            const rowCount = clearsheet.rowCount;
            for (let i = rowCount; i > 1; i--) { clearsheet.spliceRows(i, 1); }                                
            workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);  
            
            this.removeUnrelatedSheets();
        });
    }

    removeUnrelatedSheets() {
        const workbook = new ExcelJS.Workbook();
        const outputFile = (this.salesType === 'RETAIL') ? process.env.OUTPUT_FILE_ROBINSON : process.env.OUTPUT_FILE_ROBINSON_ECOMM;
        workbook.xlsx.readFile(`${outputFile}`).then(() => {
            workbook.eachSheet(sheet => {
                if (!sheet.name.startsWith('Sku_') && !sheet.name.startsWith('Store_') && sheet.name !== `${process.env.CON_SHEET_ROBINSON}`) {
                    workbook.removeWorksheet(sheet.id);
                }
            });
            // return workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE_ROBINSON}`);
            return workbook.xlsx.writeFile(`${outputFile}`);
        })
    }

    async generateOutputData() {
        try {
            const chain = this.chain;
            const salesType = this.salesType;
            const fileManager = new DataFiles();
            fileManager.source = (this.salesType === 'RETAIL') ? process.env.RAW_DATA_ROBINSON_RETAIL : process.env.RAW_DATA_ROBINSON_ECOMM;
            const files = fileManager.listFiles().filter(f => f !== `${process.env.PROCESSED}`);
            if (files.length > 0) {
                let processResult = [];
                const promises = files.map(async(file) => {
                    return await this.processGeneration(file).then((item) => {                        
                        let isCompleted = item;
                        if (isCompleted) {
                            fileManager.destination = (this.salesType === 'RETAIL') 
                                ? `${process.env.RAW_DATA_ROBINSON_RETAIL}/${process.env.PROCESSED}` 
                                : `${process.env.RAW_DATA_ROBINSON_ECOMM}/${process.env.PROCESSED}`;
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
                            statusMsg: `${chain}: ${salesType} - ${appLabels.chainMsg}`
                        }
                    }
                });

            } else {
                return {
                    isProcessed: false,
                    statusMsg: `NO DATA FILE(S) FOUND FROM ${chain}: ${salesType}!`
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
            const sourceFile = `${process.env.OUTPUT_FILE_ROBINSON}`.replace(`${process.env.OUTPUT_DIR}`, `${process.env.TEMPO_DATA_DIR}`).replace('.xlsx', '.csv');
            const sourceWB = new ExcelJS.Workbook();

            const fileManager = new DataFiles();
            fileManager.source = `${process.env.TEMPO_DATA_DIR}`;
            const fileResult = fileManager.listFiles().filter(f => f.includes(`${process.env.CON_SHEET_ROBINSON}`.toLowerCase()));

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
                                    row.getCell(22).value, // BANNER
                                    row.getCell(6).value.trim(), // BRANCH
                                    row.getCell(17).value, // SKU CATEGORY
                                    row.getCell(8).value, // DESCRIPTION
                                    parseFloat(row.getCell(11).value).toFixed(5), // PACK
                                    parseFloat(row.getCell(12).value).toFixed(5), // KG
                                    parseFloat(row.getCell(13).value).toFixed(5), // PCS
                                    parseFloat(row.getCell(14).value).toFixed(5), // GROSS
                                    parseFloat(row.getCell(16).value).toFixed(5), // NET SALES
                                    row.getCell(18).value, // AREA
                                    row.getCell(7).value, // SKU NUMBER
                                    row.getCell(23).value, // SKU PER BRAND
                                    row.getCell(24).value, // GENERALIZED SKU
                                    row.getCell(25).value, // MOTHER SKU
                                    row.getCell(26).value, // SALES CATEGORY
                                    row.getCell(27).value, // SKU DEPT
                                    row.getCell(28).value, // PLACEMENT
                                    row.getCell(29).value, // PLACEMENT REMARKS
                                    row.getCell(32).value, // CITY
                                    row.getCell(33).value, // PROVINCE
                                    row.getCell(34).value, // SKU REPORT IDENTIFIER 1
                                    row.getCell(35).value, // SKU REPORT IDENTIFIER 2
                                    "-", // SUKI CO STORE
                                    row.getCell(36).value, // INTERNAL BRAND
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

    checkFileExists(salesType, callback) {
        let attempts = 0;
        const maxAttempts = 3;
        const delay = 1000; // Delay in milliseconds between each attempt
    
        function check() {

            const outputFile = (salesType === 'RETAIL') ? process.env.OUTPUT_FILE_ROBINSON : process.env.OUTPUT_FILE_ROBINSON_ECOMM;

            fs.access(`${outputFile}`, fs.constants.F_OK, (err) => {
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

module.exports = { Robinson }