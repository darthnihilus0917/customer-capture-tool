const dotenv = require("dotenv");
dotenv.config();

const ExcelJS = require('exceljs');
const fs = require('fs');
const Papa = require('papaparse');
const { Log } = require('./logs');
const { DataFiles } = require('./files');
const { appLabels } = require('../contants/contants');
const { mergeArrays, endsWithNumber, removeLastNumber, removePrecedingString } = require('../utils/utils');

class MerryMart {
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

    async pdfToExcel() {
        try {            
            const pdfFileManager = new DataFiles();
            pdfFileManager.source = process.env.PDF_MERRYMART;
            pdfFileManager.destination = `${process.env.PDF_MERRYMART}/${process.env.PROCESSED}`;
            const pdfFiles = pdfFileManager.listFiles().filter(f => f !== `${process.env.PROCESSED}` && f.includes(`pdf`));
            if (pdfFiles.length > 0) {
                pdfFiles.forEach(async(filename) => {
                    const pdfFile = `${process.env.PDF_MERRYMART}/${filename}`;                    
                    const excelFilename = filename.replace('.pdf', '.csv');
                    const excelFile = `${process.env.CONVERTED_MERRYMART}/${excelFilename}`;
                    // convert pdf to excel
                    pdfToExcelGenerator.genXlsx(pdfFile, excelFile).then(() => {
                        if (pdfFileManager.fileExists(pdfFile)) {
                            // delete pdf after conversion
                            // pdfFileManager.deleteFile(pdfFile);
                            pdfFileManager.filename = filename;
                            pdfFileManager.moveFile();
                        }
                    }).catch((error) => {
                        console.error(error);
                        process.exit(0);
                    });
                });
                return {
                    isProcessed: true,
                    statusMsg: `${this.chain}: ${appLabels.pdfConvertion}`
                }

            } else {
                return {
                    isProcessed: false,
                    statusMsg: `NO PDF DATA FILE(S) FOUND FROM ${this.chain}!`
                } 
            }            
        } catch(error) {
            console.error(error);
            return false;
        }
    }

    async captureRawData(callback) {
        try {
            const csvFileManager = new DataFiles();
            csvFileManager.source = process.env.CONVERTED_MERRYMART;
            const csvFiles = csvFileManager.listFiles().filter(f => f !== `${process.env.PROCESSED}` && f.includes(`csv`));
            csvFiles.map((file) => {
                const csvFile = `${process.env.CONVERTED_MERRYMART}/${file}`;
                fs.readFile(csvFile, 'utf-8', (err, data) => {
                    if (err) {
                        callback(err);
                        return false;
                    }

                    const result = Papa.parse(data, { header: false });
                    const rowData = result.data;
                    let startingPoint = rowData.map((item, index) => {
                        return (item[0].length > 0 && item[0].includes(`Covering Period`)) ? index : 0;
                    }).filter(d => d !== 0);

                    const csvData = rowData.map((item, index) => {                        
                        if (index > parseInt(startingPoint) && item !== undefined) { return item.filter(val => val !== ''); }
                    }).filter(d => d !== undefined && d.length > 0);

                    // BRANCH
                    let branch = csvData.map((item, index) => { 
                        if (index === 0) { return (csvData[0].length === 1) ? csvData[1][0] : item[1] }
                    }).filter(d => d !== null)[0];
                    
                    branch = branch.toUpperCase().trim();
                    branch = (branch === 'HARBOR POINT') ? 'HARBOR POINT SUBIC' : branch;

                    // SKU DATA RANGE
                    const skuDataRange = csvData.map((item, index) => {
                        let rangeIndex = 0;
                        if (item[0].includes('ARTICLE DESCRIPTION')) { rangeIndex = index + 1; }
                        if (item[0].includes('TOTAL')) { rangeIndex = index; }
                        return rangeIndex;
                    }).filter(d => d !== 0);

                    const skuCodes = [];
                    const skuDescriptions = [];
                    let quantities = [];
                    let units = [];
                    const netSales = [];
                    const commRates = [];
                    const netPayables = [];
                    const taxClass = [];

                    const content = csvData.slice(skuDataRange[0], skuDataRange[1]);
                    content.map((item, index) => {
                        let skuCode = null;
                        let skuDesc = null;
                        let qty = null;

                        if (item[0].length >= 8 && item[0].includes('2002')) {
                            skuCode = (item[0].length > 8) ? item[0].split(' ')[0].trim() : (item[0].length === 8) ? item[0] : null;
                            skuCodes.push(parseInt(skuCode));
                        }

                        switch(branch) {
                            case "UMBRIA":
                                if (item[0].length < 3 || item[0].includes('TGM')) {
                                    if (!endsWithNumber(item[0])) {
                                        qty = item[1];
                                        if (typeof(qty) === 'string') quantities.push(qty);
                                        units.push(item[2]);
                                        netSales.push(item[3].replace(",",""));
                                        commRates.push(item[5]);
                                        netPayables.push(item[6].replace(",",""));
                                        taxClass.push(item[8]);

                                    } else {
                                        qty = removePrecedingString(item[0]);
                                        if (typeof(qty) !== 'number') quantities.push(qty);
                                        units.push(item[1]); 
                                        netSales.push(item[2].replace(",","")); 
                                        commRates.push(item[4]); 
                                        netPayables.push(item[5].replace(",",""));
                                        taxClass.push(item[7]);                                    
                                    }
                                    skuDesc = item[0].replace(/^.*?TGM/, 'TGM');
                                    skuDesc = removeLastNumber(skuDesc).trim();
                                    skuDescriptions.push(skuDesc.trim());
                                }                           
                                break;
                            default:
                                if (item[0].includes('TGM')) {
                                    if (item[0].includes('2002') && endsWithNumber(item[0])) {
                                        skuDesc = item[0].replace(/^.*?TGM/, 'TGM');
                                        skuDesc = removeLastNumber(skuDesc).trim();
                                        skuDescriptions.push(skuDesc.trim());
                                    } else if (item[0].includes('2002') && !endsWithNumber(item[0])) {
                                        skuDesc = item[0].replace(/^.*?TGM/, 'TGM');
                                        skuDescriptions.push(skuDesc.trim());
                                    } else {
                                        if (endsWithNumber(item[0])) {
                                            skuDesc = removeLastNumber(item[0]).trim();
                                        } else {
                                            skuDescriptions.push(item[0].trim());
                                        } 
                                    }                            
                                }

                                // QUANTITIES AND UNITS
                                if (item[0].length < 3 || item[0].includes('TGM') && endsWithNumber(item[0])) {
                                    if (endsWithNumber(item[0])) {
                                        qty = removePrecedingString(item[0]);
                                        if (typeof(qty) !== 'number') quantities.push(qty);
                                    }
                                }

                                if (item[0].length < 3) {
                                    const numRegex = /\d/;
                                    const letterRegex = /[a-zA-Z]/;
                                    if (letterRegex.test(item[0])) { units.push(item[0]); }
                                    if (numRegex.test(item[0])) { quantities.push(parseInt(item[0])); }
                                }   
                                
                                // NET SALES
                                if (item[1] !== undefined) { netSales.push(item[1].replace(",","")); }
                                // COMM.RATES
                                if (item[3] !== undefined) { commRates.push(item[3]); }
                                // NET PAYABLES
                                if (item[4] !== undefined) { netPayables.push(item[4].replace(",","")); }
                                // TAX CLASS
                                if (item[6] !== undefined) { taxClass.push(item[6]); }                                  
                        }
                    }).filter(d => d !== 0);

                    branch = `MM - ${branch}`;
                    quantities = quantities.filter(u => u !== undefined).filter(s => typeof(s) !== 'number' );
                    units = units.filter(u => u !== undefined).slice(0, quantities.length);

                    const skuData = mergeArrays(branch, skuCodes, skuDescriptions, quantities, units, netSales, commRates, netPayables, taxClass);
                    callback(null, skuData);
                });
            });

        } catch(err) {
            callback(err);
            return false;
        }
    }
    
    async buildRawData() {
        try {
            const chain = this.chain;
            const destinationWB = new ExcelJS.Workbook();
            const destinationFile = `${process.env.RAW_DATA_MERRYMART}/${process.env.RAW_DATA_MERRYMART_FILE}`;
            await destinationWB.xlsx.readFile(destinationFile);
            const destinationSheet = destinationWB.getWorksheet('raw');

            this.captureRawData(async(err, data) => {
                if (err) {
                    console.error(err);
                    process.exit(0);
                }
                data.forEach((item) => { destinationSheet.addRow(item) });
                await destinationWB.xlsx.writeFile(destinationFile);
            });

            const csvFileManager = new DataFiles();
            csvFileManager.source = process.env.CONVERTED_MERRYMART;
            const csvFiles = csvFileManager.listFiles().filter(f => f !== `${process.env.PROCESSED}` && f.includes(`csv`));
            csvFiles.forEach((file) => {
                const csvFile = `${process.env.CONVERTED_MERRYMART}/${file}`;
                if (csvFileManager.fileExists(csvFile)) {
                    // move file to processed
                    csvFileManager.destination = `${process.env.CONVERTED_MERRYMART}/${process.env.PROCESSED}`;
                    csvFileManager.filename = file.trim();
                    csvFileManager.moveFile();
                }
            })

            return {
                isProcessed: true,
                statusMsg: `${chain}: ${appLabels.rawDataMsg}`
            }

        } catch(e) {
            return {
                isProcessed: false,
                statusMsg: e
            }
        }
    }

    clearRawDataSheet(workbook) {
        const rawDataFile = `${process.env.RAW_DATA_MERRYMART}/${process.env.RAW_DATA_MERRYMART_FILE}`;
        workbook.xlsx.readFile(rawDataFile).then(() => {
            const clearsheet = workbook.getWorksheet(`${process.env.RAW_DATA_MERRYMART_SHEET}`);
            const rowCount = clearsheet.rowCount;
            for (let i = rowCount; i > 1; i--) { clearsheet.spliceRows(i, 1); }                                
            workbook.xlsx.writeFile(rawDataFile);  
        });
    }

    async processGeneration(filename) {
        try {
            const currentDate = new Date();

            const sourceFile = `${process.env.RAW_DATA_MERRYMART}/${filename}`;
            const sourceSheetName = `${process.env.RAW_DATA_MERRYMART_SHEET}`;
            const sourceWB = new ExcelJS.Workbook();

            return await sourceWB.xlsx.readFile(sourceFile).then(() => {
                const sourceSheet = sourceWB.getWorksheet(sourceSheetName);

                const destinationWB = new ExcelJS.Workbook();
                destinationWB.xlsx.readFile(`${process.env.OUTPUT_FILE}`).then(async() => {
                    const destinationSheet = destinationWB.getWorksheet(`${process.env.CON_SHEET_MERRYMART}`);

                    const showcaseSheet = destinationWB.getWorksheet(`${process.env.STORE_SHEET_SHOWCASE}`);
                    const srpSheet = destinationWB.getWorksheet(`${process.env.STORE_SHEET_SRP}`);
                    const vamSheet = destinationWB.getWorksheet(`${process.env.STORE_SHEET_VAM}`);

                    const consolidatedSheet = destinationWB.getWorksheet(`${process.env.SKU_SHEET_CONSOLIDATED}`);
                    const commrateSheet = destinationWB.getWorksheet(`${process.env.SKU_SHEET_COMMRATE}`);
                    const ninersSheet = destinationWB.getWorksheet(`${process.env.SKU_SHEET_NINERS}`);

                    sourceSheet.eachRow({ includeEmpty: false, firstRow: 2 }, (row, rowNumber) => {
                        const rowData = [1, 2, 3, 4, 5, 6, 7, 8, 9].map(col => row.getCell(col).value);
                        if (rowNumber > 1) {
                            const cutOffSegments = this.cutOff.split(' ');
                            const cutOffValue = this.cutOff;

                            const newRowData = [
                                currentDate.getFullYear(), // YEAR
                                cutOffSegments[0].toUpperCase(), // MONTH
                                this.cutOff, // CUT OFF 
                                rowData[8], // BRANCH
                                parseInt(rowData[0]), // ARTICLE
                                rowData[1], // ARTICLE DESCRIPTION
                                parseFloat(rowData[2]).toFixed(5), // ORIG QTY
                                rowData[3], // SALES UNIT
                                parseFloat(0).toFixed(5), // PACK
                                parseFloat(0).toFixed(5), // KG
                                parseFloat(0).toFixed(5), // PCS
                                parseFloat(rowData[4]).toFixed(5), // GROSS SALES
                                rowData[5], // COMM RATE
                                parseFloat(rowData[6]).toFixed(5), // NET PAYABLE
                                rowData[7], // TAX CLASS
                                "-", // SKU CATEGORY
                                "-", // AREA
                                "-", // KAM
                                "-", // CHAIN
                                "-", // BANNER
                                "-", // SKU PER BRAND
                                "-", // GENERALIZED SKU
                                "-", // MOTHER SKU
                                "RETAIL", // SALES CATEGORY
                                "-", // SKU DEPT.
                                "-", // PLACEMENT
                                "-", // PLACEMENT REMARKS
                            ]
                            destinationSheet.addRow(newRowData);
                        }
                    });
                    await destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);

                    destinationSheet.eachRow({ includeEmpty: false, firstRow: 2}, (row, rowNumber) => {
                        if (rowNumber > 1) {
                            row.getCell(4).alignment = { horizontal: 'left' }; // BRANCH
                            row.getCell(5).alignment = { horizontal: 'center' }; // ARTICLE
                            row.getCell(6).alignment = { horizontal: 'left' }; // ARTICLE DESC
                            row.getCell(7).alignment = { horizontal: 'right' }; // ORIG QTY
                            row.getCell(8).alignment = { horizontal: 'center' }; // SALES UNIT
                            row.getCell(9).alignment = { horizontal: 'right' }; // PACK
                            row.getCell(9).numFmt = `###0.00000`;
                            row.getCell(9).value = { formula: `IF(H${rowNumber}="PC", G${rowNumber}, 0)`};
                            row.getCell(10).alignment = { horizontal: 'right' }; // KG
                            row.getCell(10).numFmt = `###0.00000`;
                            row.getCell(10).value = { formula: `IF(H${rowNumber}="KG", G${rowNumber}, G${rowNumber} * VLOOKUP(E${rowNumber},Sku_Consolidated!A2:U${consolidatedSheet.lastRow.number},{21},FALSE))`};
                            row.getCell(11).alignment = { horizontal: 'right' }; // PCS
                            row.getCell(12).alignment = { horizontal: 'right' }; // GROSS SALES
                            row.getCell(12).numFmt = `###0.00000`;
                            row.getCell(13).alignment = { horizontal: 'center' }; // COMM RATE
                            row.getCell(14).alignment = { horizontal: 'right' }; // NET PAYABLE
                            // TAX CLASS
                            row.getCell(16).value = { formula: `VLOOKUP(E${rowNumber},Sku_Consolidated!A2:G${consolidatedSheet.lastRow.number}, 7, FALSE)`}; // SKU CATEGORY
                            row.getCell(17).value = { formula: `IF(IFERROR(VLOOKUP(D${rowNumber},Store_Showcase!C2:H${showcaseSheet.lastRow.number},{6}, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(D${rowNumber},Store_SRP!C2:H${srpSheet.lastRow.number},{6}, FALSE), TRUE)=TRUE,VLOOKUP(C${rowNumber},Store_VAM!C2:H${vamSheet.lastRow.number},{6}, FALSE),VLOOKUP(D${rowNumber},Store_SRP!C2:H${srpSheet.lastRow.number},{6}, FALSE)), VLOOKUP(D${rowNumber},Store_Showcase!C2:H${showcaseSheet.lastRow.number},{6}, FALSE))`}; // AREA
                            row.getCell(18).value = { formula: `IF(IFERROR(VLOOKUP(D${rowNumber},Store_Showcase!C2:I${showcaseSheet.lastRow.number},{7}, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(D${rowNumber},Store_SRP!C2:I${srpSheet.lastRow.number},{7}, FALSE), TRUE)=TRUE,VLOOKUP(D${rowNumber},Store_VAM!C2:I${vamSheet.lastRow.number},{7}, FALSE),VLOOKUP(D${rowNumber},Store_SRP!C2:I${srpSheet.lastRow.number},{7}, FALSE)), VLOOKUP(D${rowNumber},Store_Showcase!C2:I${showcaseSheet.lastRow.number},{7}, FALSE))`}; // KAM
                            row.getCell(19).value = { formula: `IF(IFERROR(VLOOKUP(D${rowNumber},Store_Showcase!C2:K${showcaseSheet.lastRow.number},{9}, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(D${rowNumber},Store_SRP!C2:K${srpSheet.lastRow.number},{9}, FALSE), TRUE)=TRUE,VLOOKUP(D${rowNumber},Store_VAM!C2:K${vamSheet.lastRow.number},{9}, FALSE),VLOOKUP(D${rowNumber},Store_SRP!C2:K${srpSheet.lastRow.number},{9}, FALSE)), VLOOKUP(D${rowNumber},Store_Showcase!C2:K${showcaseSheet.lastRow.number},{9}, FALSE))`}; // CHAIN
                            row.getCell(20).value = { formula: `IF(VLOOKUP(E${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)="SHOWCASE", VLOOKUP(D${rowNumber},Store_Showcase!C2:L${showcaseSheet.lastRow.number},{10}, FALSE), IF(VLOOKUP(E${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)="VAM",VLOOKUP(D${rowNumber},Store_VAM!C2:L${vamSheet.lastRow.number},{10}, FALSE),VLOOKUP(D${rowNumber},Store_SRP!C2:L${srpSheet.lastRow.number},{10}, FALSE)))`}; // BANNER
                            row.getCell(21).value = { formula: `VLOOKUP(E${rowNumber},Sku_Consolidated!A2:K${consolidatedSheet.lastRow.number}, 11, FALSE)`}; // SKU PER BRAND
                            row.getCell(22).value = { formula: `VLOOKUP(E${rowNumber},Sku_Consolidated!A2:H${consolidatedSheet.lastRow.number}, 8, FALSE)`}; // GENERALIZED SKU
                            row.getCell(23).value = { formula: `VLOOKUP(E${rowNumber},Sku_Consolidated!A2:I${consolidatedSheet.lastRow.number}, 9, FALSE)`}; // MOTHER SKU
                            // SALES CATEGORY
                            row.getCell(25).value = { formula: `VLOOKUP(E${rowNumber},Sku_Consolidated!A2:E${consolidatedSheet.lastRow.number}, 5, FALSE)`}; // SKU DEPT.
                            row.getCell(26).value = { formula: `IF(VLOOKUP(E${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)="SHOWCASE", VLOOKUP(D${rowNumber},Store_Showcase!C2:N${showcaseSheet.lastRow.number},{12}, FALSE), IF(VLOOKUP(E${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)="VAM",VLOOKUP(D${rowNumber},Store_VAM!C2:N${vamSheet.lastRow.number},{12}, FALSE),VLOOKUP(D${rowNumber},Store_SRP!C2:N${srpSheet.lastRow.number},{12}, FALSE)))`}; // PLACEMENT
                            row.getCell(27).value = { formula: `IF(IFERROR(Z${rowNumber},TRUE)=TRUE, "-","OK")`}; // PLACEMENT REMARKS
                            row.getCell(28).value = { formula: `IF(IFERROR(VLOOKUP(D${rowNumber},Store_Showcase!C2:H${showcaseSheet.lastRow.number},{4}, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(D${rowNumber},Store_SRP!C2:H${srpSheet.lastRow.number},{4}, FALSE), TRUE)=TRUE,VLOOKUP(C${rowNumber},Store_VAM!C2:H${vamSheet.lastRow.number},{4}, FALSE),VLOOKUP(D${rowNumber},Store_SRP!C2:H${srpSheet.lastRow.number},{4}, FALSE)), VLOOKUP(D${rowNumber},Store_Showcase!C2:H${showcaseSheet.lastRow.number},{4}, FALSE))`}; // CITY
                            row.getCell(29).value = { formula: `IF(IFERROR(VLOOKUP(D${rowNumber},Store_Showcase!C2:H${showcaseSheet.lastRow.number},{5}, FALSE), TRUE)=TRUE, IF(IFERROR(VLOOKUP(D${rowNumber},Store_SRP!C2:H${srpSheet.lastRow.number},{5}, FALSE), TRUE)=TRUE,VLOOKUP(C${rowNumber},Store_VAM!C2:H${vamSheet.lastRow.number},{5}, FALSE),VLOOKUP(D${rowNumber},Store_SRP!C2:H${srpSheet.lastRow.number},{5}, FALSE)), VLOOKUP(D${rowNumber},Store_Showcase!C2:H${showcaseSheet.lastRow.number},{5}, FALSE))`}; // PROVINCE
                            row.getCell(30).value = { formula: `VLOOKUP(E${rowNumber},Sku_Consolidated!A2:R${consolidatedSheet.lastRow.number},18,FALSE)`}; // SKU IDENTIFIER 1
                            row.getCell(31).value = { formula: `VLOOKUP(E${rowNumber},Sku_Consolidated!A2:S${consolidatedSheet.lastRow.number},19,FALSE)`}; // SKU IDENTIFIER 2
                            row.getCell(32).value = { formula: `VLOOKUP(E${rowNumber},Sku_Consolidated!A2:T${consolidatedSheet.lastRow.number},20,FALSE)`}; // INTERNAL BRAND

                        }
                    });

                    destinationWB.xlsx.writeFile(`${process.env.OUTPUT_FILE}`).then(() => {
                        const fileManager = new DataFiles();
                        fileManager.copyFile(`${process.env.OUTPUT_FILE}`,`${process.env.OUTPUT_FILE_MERRYMART}`);

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
            const clearsheet = workbook.getWorksheet(`${process.env.CON_SHEET_MERRYMART}`);
            const rowCount = clearsheet.rowCount;
            for (let i = rowCount; i > 1; i--) { clearsheet.spliceRows(i, 1); }                                
            workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE}`);  
            
            this.removeUnrelatedSheets();
        });
    }

    removeUnrelatedSheets() {
        const workbook = new ExcelJS.Workbook();
        workbook.xlsx.readFile(`${process.env.OUTPUT_FILE_MERRYMART}`).then(() => {
            workbook.eachSheet(sheet => {
                if (!sheet.name.startsWith('Sku_') && !sheet.name.startsWith('Store_') && sheet.name !== `${process.env.CON_SHEET_MERRYMART}`) {
                    workbook.removeWorksheet(sheet.id);
                }
            });
            return workbook.xlsx.writeFile(`${process.env.OUTPUT_FILE_MERRYMART}`);
        })
    }       

    generateOutputData() {
        try {
            const chain = this.chain;
            const fileManager = new DataFiles();
            fileManager.source = process.env.RAW_DATA_MERRYMART;
            const files = fileManager.listFiles().filter(f => f.includes('xlsx'));
            if (files.length > 0) {
                let processResult = [];
                const promises = files.map(async(file) => {
                    return await this.processGeneration(file).then((item) => {
                        let isCompleted = item;
                        if (isCompleted) {
                            const workbook = new ExcelJS.Workbook();
                            this.clearRawDataSheet(workbook);
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
            const sourceFile = `${process.env.OUTPUT_FILE_MERRYMART}`.replace(`${process.env.OUTPUT_DIR}`, `${process.env.TEMPO_DATA_DIR}`).replace('.xlsx', '.csv');
            const sourceWB = new ExcelJS.Workbook();

            const fileManager = new DataFiles();
            fileManager.source = `${process.env.TEMPO_DATA_DIR}`;
            const fileResult = fileManager.listFiles().filter(f => f.includes(`${process.env.CON_SHEET_MERRYMART}`.toLowerCase()));

            if (fileResult.length > 0) {
                return await sourceWB.csv.readFile(sourceFile, { encoding: 'utf-8' }).then(() => {
                    const sourceSheet = sourceWB.worksheets[0];

                    const destinationWB = new ExcelJS.Workbook();
                    destinationWB.xlsx.readFile(`${process.env.CONSOLIDATED_DATA_FILE}`).then(async() => {
                        const destinationSheet = destinationWB.getWorksheet(`${process.env.CONSOLIDATED_SHEET}`);

                        sourceSheet.eachRow({ includeEmpty: false, firstRow: 2 }, (row, rowNumber) => {
                            if (rowNumber > 1 && rowNumber !== undefined) {

                                let city = (row.getCell(28).value.includes(`�`))
                                    ? row.getCell(28).value.replace(`�`, 'ñ') 
                                    : row.getCell(28).value

                                const newRowData = [
                                    row.getCell(1).value, // YEAR
                                    row.getCell(2).value, // MONTH
                                    row.getCell(19).value, // CHAIN
                                    row.getCell(20).value, // BANNER
                                    row.getCell(4).value.trim(), // BRANCH
                                    row.getCell(16).value, // SKU CATEGORY
                                    row.getCell(6).value, // DESCRIPTION
                                    parseFloat(row.getCell(9).value).toFixed(5), // PACK
                                    parseFloat(row.getCell(10).value).toFixed(5), // KG
                                    parseFloat(row.getCell(11).value).toFixed(5), // PCS
                                    parseFloat(row.getCell(12).value).toFixed(5), // GROSS
                                    parseFloat(row.getCell(14).value).toFixed(5), // NET SALES
                                    row.getCell(17).value, // AREA
                                    row.getCell(5).value, // SKU NUMBER
                                    row.getCell(21).value, // SKU PER BRAND
                                    row.getCell(22).value, // GENERALIZED SKU
                                    row.getCell(23).value, // MOTHER SKU
                                    row.getCell(24).value, // SALES CATEGORY
                                    row.getCell(25).value, // SKU DEPT
                                    row.getCell(26).value, // PLACEMENT
                                    row.getCell(27).value, // PLACEMENT REMARKS
                                    // row.getCell(28).value.toString().trim(), // CITY
                                    city, // CITY
                                    row.getCell(29).value, // PROVINCE
                                    row.getCell(30).value, // SKU REPORT IDENTIFIER 1
                                    row.getCell(31).value, // SKU REPORT IDENTIFIER 2
                                    "-", // SUKI CO STORE
                                    row.getCell(32).value, // INTERNAL BRAND
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
            fs.access(`${process.env.OUTPUT_FILE_MERRYMART}`, fs.constants.F_OK, (err) => {
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

module.exports = { MerryMart }