const xlsx = require('node-xlsx');
const fs = require('fs');
const path = require('path');
let rows = []
const directoryPath = `${__dirname}excel\\`;

async function parseDirectory() {
    return new Promise((resolve, reject) => fs.readdir(directoryPath, function(err, files) {
        if (err){
            return console.log('Unable to scan directory: ' + err);
            reject();
        }
        files.forEach(function(file) {
            if (file.includes('aduri')) {
                let obj = xlsx.parse(`${__dirname}excel\\${file}`);
                for (let i = 0; i < obj.length; i++) {
                    let sheet = obj[i];
                    for (let j = 0; j < sheet['data'].length; j++) {
                        rows.push(sheet['data'][j]);
                    }
                }
            }
        });
        resolve();
    }))
}

async function main() {
    await parseDirectory();

    const buffer = xlsx.build([{
        name: "aduri",
        data: rows
    }])

    fs.writeFile(`${__dirname}/aduri.xlsx`, buffer, function(err) {
        if (err) {
            return console.log(err);
        }
        console.log("aduri.csv was saved in the current directory!");
    });
}
main();