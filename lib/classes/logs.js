const fs = require('fs');

class Log {
    constructor() {
        this.meat = null;
        this.filePath = null;
        this.action = null;
    }

    setChain(meat) { this.meat = meat; }
    getChain() { return this.meat; }

    setFilePath(filePath) { this.filePath = filePath; }
    getFilePath() { return this.filePath; }

    setAction(action) { this.action = action; }
    getAction() { return this.action; }

    logActivity() {
        const currentDate = new Date();
        const logDate = `${currentDate.toLocaleDateString()} ${currentDate.toLocaleTimeString()}`;
        let log = `${logDate} - ${this.meat} - ${this.action} - processed`;

        fs.readFile(this.filePath, 'utf8', (err, data) => {
            if (err) {
                console.error(err);
                return;
            }

            const modifiedData = data.endsWith('\n') ? data + log : data + '\n' + log;

            fs.writeFile(this.filePath, modifiedData, 'utf8', (err) => {
                if (err) {
                    console.error(err);
                    return;
                }
            })
        })
    }
}

module.exports = { Log }