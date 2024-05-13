const appLabels = {
    confirmMeatProcessing: 'Do you still want to process a Meat Data Source? (yes/no)',
    confirmProcessing: 'Do you still want to process a store? (yes/no)',
    confirmExit: "Are you sure you want to exit? (yes/no)",
    invalidAnswer: "Invalid answer. Please select a number from the provided options.",
    closingApp: 'Closing OTC Customer Capture Tool. Goodbye!',
    dataSourceMsg: 'Meat Data Source generation complete! Please check the output folder.',
    sotcDataMsg: "SOTC & Pickup data copied! You may begin generating data source anytime.",
    consolidationMsg: "Data consolidation complete! Please check the final folder.",
    processNotAvailable: "This process is not available to",
    pdfConvertion: "PDF conversion complete! Please check the raw data converted folder.",
    noSapFile: "No SAP exported file found!",
    tooManyFiles: "Too many SOTC files to process. Only one file is allowed in the folder."
}

module.exports = { appLabels }