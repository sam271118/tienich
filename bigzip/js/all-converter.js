

import {
    Archive
} from './node_modules/libarchive.js/main.js';
Archive.init({
    workerUrl: '../../js/node_modules/libarchive.js/dist/worker-bundle.js'
});

window.onExtractFiles = function onExtractFiles(file) {
    return new Promise(function (resolve, reject) {
        Archive.open(file).then(async function (archive) {
            const result = await archive.extractFiles();
            resolve(result);
        }).catch(function (error) {
            reject(error);
        });
    })
};