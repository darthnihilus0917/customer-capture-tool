const path = require('path');
const fsFolders = require('fs-extra');
const fs = require('fs');

const destinationBuildFolder = "./build";

const folderList = [
    'final',
    'output',
    'raw_data',
    'templates',
    'monitoring'
]

const files = [
    'logs.txt',
    '.env'
]

const deployFiles = () => {
    // COPY FILE
    files.forEach((file) => {
        const fileToBeCopied = path.basename(file);    
        const destination = path.join(destinationBuildFolder, fileToBeCopied);
        fs.copyFileSync(file, destination);
    });

    // CREATE FOLDER
    folderList.forEach((folder) => {
        const folderPath = path.basename(folder)
        const newFolderPath = path.join(destinationBuildFolder, folderPath);
        
        fs.mkdir(newFolderPath, { recursive: true }, (err) => {
            if (err) {
                console.error(`Error creating folder: ${err}`);
                process.exit(0);
            }      
        });

        // COPY SUBFOLDER CONTENT
        const sourceFolder = `./${folder}`;
        fsFolders.copy(sourceFolder, `${destinationBuildFolder}/${folder}`, (err) => {
            if (err) {
                console.error(`Error copying files: ${err}`);
                process.exit(0);
            }  
        });
    });
}

try {
    if (fs.existsSync(destinationBuildFolder)) {
        deployFiles();
    } else {
        // CREATE BUILD DIRECTORY
        fs.mkdirSync(destinationBuildFolder);
        // CREATE FOLDERS AND COPY FILES
        deployFiles();
    }

} catch (err) {
    console.error(`Error checking BUILD folder existence: ${err}`);
}
