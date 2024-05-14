const fs = require('fs');

const loader = () => {
    const frames = ['-', '\\', '|', '/'];
    let currentFrameIndex = 0;

    function updateLoader() {
        process.stdout.write(`\rProcessing ${frames[currentFrameIndex]}`);
        currentFrameIndex = (currentFrameIndex + 1) % frames.length;
    }

    const intervalId = setInterval(updateLoader, 100);

    return function stopLoader() {
        clearInterval(intervalId);
        process.stdout.write('\r');
    };
}

const startsWithZero = (value) => {
    return value.startsWith('0');
}

const removeLeadingZero = (value) => {
    return value.replace(/^0+/, '');
}

const convertPath = (path) => {
    path = path.replace(/\\/g, '/');
    return path;
}

const cutOffFormat = (value) => {
    const regex = /^([A-Za-z]{1,5}|[A-Za-z]{5}\s\d{1,2})\s\d{1,2}\s(TO|-)\s\d{1,2}$/i;
    return regex.test(value);
}

const loadTitle = () => {
    console.log("\n===============================");
    console.log("CUSTOMER CAPTURE TOOL v0.0.1");
    console.log("===============================");
}

const rawDataDateFormat = (dateValue) => {
    const day = new Date(dateValue).getDate().toLocaleString();
    const month = new Date(dateValue).getMonth() + 1;
    const year = new Date(dateValue).getFullYear();
    return `${month}/${day}/${year}`;
}

const mergeArrays = (branch, ...arrays) => {
    const length = arrays.reduce((minLength, arr) => Math.min(minLength, arr.length), Infinity);
    const merged = [];

    for (let i = 0; i < length; i++) {
        const newArray = arrays.map(arr => arr[i]);
        newArray.push(branch);
        merged.push(newArray);
    }    
    return merged;
}

const endsWithNumber = (str) => {
    const regex = /\d$/;
    return regex.test(str);
}

const removeLastNumber = (str) => {
    return str.replace(/\d$/, '');
}

const removePrecedingString = (str) => {
    const match = str.match(/\d+$/);
    return match ? match[0] : '';
}

const isNumeric = (value) => {
    const numRegex = /\d/;
    return numRegex.test(value);
}

const cleanNegativeValue = (value) => {
    let match = null;
    if (value.includes("(")) {
        match = value.match(/\(([^)]+)\)/);
        match = (match[1].includes('-')) ? match[1].match(/^([^\-]+)\-/)[1] : match[1];
        match = parseFloat(match.replace(",","")) * -1;
        return match.toFixed(5);
        
    } else if (value.includes("-")) {
        match = value.match(/^([^\-]+)\-/)[1]
        match = parseFloat(match.replace(",","")) * -1;
        return match.toFixed(5);
    }
    return parseFloat(value.replace(",","")).toFixed(5);    
}

const pickupHypen = (value) => {
    if (value.includes('PICK UP')) {
        return value.replace(/(PICK UP)(?! - )/g, 'PICK UP -');
    }
    return value;
}

module.exports = { 
    loader, 
    loadTitle,
    startsWithZero,
    removeLeadingZero,
    convertPath,
    cutOffFormat,
    rawDataDateFormat,
    mergeArrays,
    endsWithNumber,
    removeLastNumber,
    removePrecedingString,
    isNumeric,
    cleanNegativeValue,
    pickupHypen
}