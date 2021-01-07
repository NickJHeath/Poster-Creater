const exec = require('child_process').exec;
const remote = require('electron').remote;
const dialog = remote.dialog;
const fs = require('fs');
const path = require('path');
const find = require('find');
const which = require('which');
var quality = '400';
var LOloc = 'C:\\Program Files\\LibreOffice\\program\\soffice.bin';

/**
 * List files in a folder using an asynchronous function.
 * @param cmd {string}
 * @return {Promise<string>}
 Written by self but modified structure from https://medium.com/@ali.dev/how-to-use-promise-with-exec-in-node-js-a39c4d7bbf77
 */

function listFilesAsync(folder) {
    return new Promise((resolve, reject) => {
        fs.readdir(folder, (err, files) => {
            if (err) {
                console.warn(err);
            }
            resolve(files ? files : err);
        });
    });
}

/**
 * Executes a shell command and return it as a Promise.
 * @param cmd {string}
 * @return {Promise<string>}
 https://medium.com/@ali.dev/how-to-use-promise-with-exec-in-node-js-a39c4d7bbf77
 */

function execShellCommand(cmd) {
    const exec = require('child_process').exec;
    return new Promise((resolve, reject) => {
        exec(cmd, (error, stdout, stderr) => {
            if (error) {
                console.warn(error);
            }
            resolve(stdout ? stdout : stderr);
        });
    });
}

/**
 * Asynchronously check whether file exists with timer to allow for file creation.
 * @param filePath {string}
 * @param timeout {int>}
 * @return {Promise<string>}
    https://stackoverflow.com/questions/26165725/nodejs-check-file-exists-if-not-wait-till-it-exist
 */
function checkExistsWithTimeout(filePath, timeout) {
    return new Promise(function (resolve, reject) {

        var timer = setTimeout(function () {
            watcher.close();
            reject(new Error('File did not exists and was not created during the timeout.'));
        }, timeout);

        fs.access(filePath, fs.constants.R_OK, function (err) {
            if (!err) {
                clearTimeout(timer);
                watcher.close();
                resolve();
            }
        });

        var dir = path.dirname(filePath);
        var basename = path.basename(filePath);
        var watcher = fs.watch(dir, function (eventType, filename) {
            if (eventType === 'rename' && filename === basename) {
                clearTimeout(timer);
                watcher.close();
                resolve();
            }
        });
    });
}

/**
 * Check PDF has been created then transform first page of PDF to JPG
 * @param files {obj}
 * @return void
 */
async function pdfToJpg(files) {
    for (var file of files) {
        await checkExistsWithTimeout(file['in'], 600000);
        var command = 'magick -density ' + quality + ' "' + file['in'] + '[0]" ' + '"' + file['out'] + '"';
        await execShellCommand(command);
    }
    dialog.showMessageBoxSync(null, {
        title: "Success",
        message: "Posters created!",
        buttons: ["OK"]
    });
};

/**
 * Convert PowerPoint presentation to PDF 
 * @param files {obj}
 * @return void
 */
function pptxToPDF(files) {
    return new Promise(async (resolve, reject) => {
        for (var file of files) {
            var command = '"' + LOloc.replace(/\\/g, '\\') + '" --convert-to pdf --outdir "' + file['in'] + '"';
            await execShellCommand(command);
        }
        resolve('Done');
    });
};

/**
 * Delete all PDFs created during conversion process after first checking those PDFs exist
 * @param files {obj}
 * @return void
 */
async function deletePDFs(files) {
    for (var file of files) {
        await checkExistsWithTimeout(file['out'], 600000);
        var command = 'del "' + file['in'] + '"';
        await execShellCommand(command);
    }
}

/**
 * Takes list of PowerPoint files in folder and uses it to generate filenames for PDFs and JPGs
 * @param files {obj}
 * @param folder {string}
 * @param type {array <string>}
 * @return list {obj}
 */
function getList(files, folder, type) {
    if (type === '.pdf') {
        formattedFolder = folder.replace(/\\/g, '\\') + '" ' + '"' + folder.replace(/\\/g, '\\') + "\\";
    }
    else {
        formattedFolder = folder.replace(/\\/g, '\\') + '\\';
    }
    var pptxList = files.filter(file => file.match(/\.[0-9a-z]+$/i) && file.match(/\.[0-9a-z]+$/i)[0] === '.pptx');
    var list = pptxList.map(file => {
        var pdf = file.replace(/\.pptx/g, '.pdf');
        var jpg = file.replace(/\.pptx/g, '.jpg');
        if (type === '.pdf')
            return { 'in': formattedFolder + file, 'out': formattedFolder + pdf };
        else
            return { 'in': formattedFolder + pdf, 'out': formattedFolder + "posters\\" + jpg };
    });
    return list;
}

/**
 * Gets a list of files from the folder chosen by the user
 * @return void
 */
async function chooseFiles() {
    var options = {
        title: 'Simple dialog',
        properties: ['openFile', 'openDirectory'],
    };
    var selection = await dialog.showOpenDialog(options);
    if (selection.canceled === true) {
        dialog.showMessageBoxSync(null, {
            title: "Folder prompt",
            message: "Please choose a folder",
            buttons: ["OK"]
        });
    }
    return selection;
}

/**
 * Allow the user to select the soffice.bin file needed to convert PowerPoints to PDFs, only called if not in default location
* @return void
*/
async function findLO() {
    var options = {
        properties: ['openFile'],
        filters: [{ name: 'All Files', extensions: ['*'] }],
        title: "Please choose the soffice.bin file",
        message: "Please select soffice.bin within LibreOffice install directory (usually within Program Files directory)"
    };
    dialog.showMessageBoxSync(null, {
        title: "Folder prompt",
        message: "Please select the soffice.bin file from within your LibreOffice install directory (file is usually found within LibreOffice\\Program directory)",
        buttons: ["OK"]
    });
    var selection = await dialog.showOpenDialog(options);
    if (selection.canceled === true) {
        dialog.showMessageBoxSync(null, {
            title: "File prompt",
            message: "Please choose the soffice.bin file",
            buttons: ["OK"]
        });
    }
    return selection.filePaths[0];
}

/**
 * Shows message box telling user that poster conversion process has begun. 
* @return void
*/
function convertMsg() {
    dialog.showMessageBoxSync(null, {
        title: "Converting PowerPoint presentations",
        message: "Posters being created",
        buttons: ["OK"]
    });
}

/**
* Get the files from the folder chosen by the user
* @param selection {obj} 
* @return void
*/
async function getFiles(selection) {
    var folder = selection.filePaths[0];
    var files = await listFilesAsync(folder);
    return files;
}

/**
 * Build list of filenames for the PDFs and JPGs that will be generated by conversion process. 
 * @param files {obj}
 * @param folder {string}
 * @return list {obj}
 */
function buildList(files, folder) {
    var fileList = ['.pdf', '.jpg'];
    var list = [];
    for (var type of fileList) {
        list.push(getList(files, folder, type));
    }
    return list;
}

/**
 * Chain together the functions needed to convert the PowerPoints presentations to JPGs. 
 * @param files {obj}
 * @return void
 */
async function pptxToJPG(files) {
    convertMsg();
    pptxToPDF(files[0]);
    pdfToJpg(files[1]);
    deletePDFs(files[1]);
}

// Begin the conversion process after user selects folder containing PowerPoint presentations
if ($("#getFolder").length) {
    $("#getFolder").click(async () => {
        var selection = await chooseFiles();
        if (selection.canceled === true) {
            return;
        }

        // check if LibreOffice soffice.bin file is in the expected location, if not ask user to specify its location
        if (!(fs.existsSync('C:\\Program Files\\LibreOffice\\program\\soffice.bin')))
            LOloc = await findLO();
        var files = await getFiles(selection);
        var dir = selection.filePaths[0];
        var finalFiles = buildList(files, dir);
        await execShellCommand('mkdir ' + dir + '\\posters');
        pptxToJPG(finalFiles);
        return false;
    });
}

// Set HD quality for the poster
if ($("#setHD").length) {
    $("#setHD").click(() => {
        quality = '160';
        return false;
    });
}

// Set Ultra HD quality for the poster
if ($("#setUHD").length) {
    $("#setUHD").click(() => {
        quality = '400';
        return false;
    });
}