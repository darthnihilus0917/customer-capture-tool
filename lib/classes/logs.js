const fs = require('fs');

class Log {
    constructor() {
        this.chain = null;
        this.filePath = null;
        this.action = null;
    }

    setChain(chain) { this.chain = chain; }
    getChain() { return this.chain; }

    setSalesType(salesType) { this.salesType = salesType; }
    getSalesType() { return this.salesType; }

    setFilePath(filePath) { this.filePath = filePath; }
    getFilePath() { return this.filePath; }

    setAction(action) { this.action = action; }
    getAction() { return this.action; }

    logActivity() {
        const currentDate = new Date();
        const logDate = `${currentDate.toLocaleDateString()} ${currentDate.toLocaleTimeString()}`;
        let log = (this.chain === 'ROBINSON')
            ? `${logDate} - ${this.chain}: ${this.salesType} - ${this.action} - processed`
            : `${logDate} - ${this.chain} - ${this.action} - processed`;

        fs.readFile(this.filePath, 'utf8', (err, data) => {
            if (err) {
                console.error(err);
                return;
            }

            const modifiedData = data.endsWith('\n') ? data + log : data + '\n' + log;
            ;

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