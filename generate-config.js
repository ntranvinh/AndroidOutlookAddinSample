#!/usr/bin/env node

let fs = require('fs');

if (process.argv.length === 3) {
    switch (process.argv[2]) {
        case 'localhost':
            copyConfig('localhost')
            break
        case 'production':
            copyConfig('production')
            break
        default:
            break
    }
}

function copyConfig(name) {
    let path = 'config-template/config-' + name + '.js'
    fs.readFile(path, function (err, data) {
        fs.writeFileSync('src/config/index.js', data)
    });
    console.log(path + ' has been copied to src/config/index.js')
}