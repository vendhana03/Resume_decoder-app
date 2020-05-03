/**
 * Framework imports
 */
const fs = require("fs");
const pdf = require("pdf-parse");
const excel = require("exceljs");
var mongoose = require("mongoose");
 
/**
 * Feature imports
 */
const constFile = require("./constant");
var ResumeData = require('../resume decoder/schema');

/**
 * Global variables
 */
let worksheet, workbook1, filesList;

/**
 * Reading the directory
 */
try {
    filesList = fs.readdirSync(constFile.srcFolderPath);
} catch (error) {
    throw new Error(error);
}

/**
 * To check the storage location from constant file
 */
if(constFile.storageLocation == constFile.strFile){

    /**
     * Initialize a workbook
     */
    workbookInit();
    
    /**
     * Creating array of promises for the fileslist to store in files
     */
    let requests = filesList.map((file) => {
        return new Promise((resolve) => {
        dataExtractorAndFileStorage(file, resolve);
        });
    });

    /**
     * Function after all promises for fileslist are resolved
     */
    Promise.all(requests).then(() => {
        return createWorkBook();
    }).catch((error) => {
        throw new Error(error);
    });

}else if(constFile.storageLocation == constFile.strDb) {

    /**
     * Connecting to MongoDB
     */
    mongoose.connect(constFile.dbUrl, { useNewUrlParser: true , useCreateIndex:true, useUnifiedTopology: true}, function (err) {
        if (err) throw err;
        console.log(constFile.strConnectedSuccess);

        /**
         * Creating array of promises for the fileslist to store in DB
         */
        let requests = filesList.map((file) => {
            return new Promise((resolve) => {
                dataExtractorAndDBStorage(file, resolve);
            });
        });

        /**
         * Function after all promises for fileslist are resolved
         */
        Promise.all(requests).then(() => {
            console.log(constFile.strDataAdded);
            process.exit();
        }).catch((error) => {
            throw new Error(error);
        });
    });
}else{
    throw new Error(constFile.strInvalidLocation);
}

/**
 * Function to initialize a workbook with heders and styles.
 */
function workbookInit(){

    try {
        
        /**
         * To initialize a new excel workbook
         */
        workbook1 = new excel.Workbook();
        workbook1.creator = constFile.workbookCreator;
        workbook1.lastModifiedBy = constFile.workbookLastModifiedBy;
        workbook1.created = new Date();
        workbook1.modified = new Date();

        /**
         * Adding a new worksheet
         */
        const sheet1 = workbook1.addWorksheet(constFile.worksheetName);
        worksheet = workbook1.getWorksheet(constFile.worksheetName);

        /**
         * Initializing the column headers for the worksheet
         */
        const reColumns=[
            {header: constFile.headerFileName, key: constFile.keyFileName, width:35},
            {header: constFile.headerEmailID, key: constFile.keyEmailID, width:40},
            {header: constFile.headerAltEmails, key: constFile.keyAltEmails, width:40},
            {header: constFile.headerContact, key: constFile.keyContact, width:30},
            {header: constFile.headerAltContacts, key: constFile.keyAltContacts, width:30}
        ];
        sheet1.columns = reColumns;

        /**
         * Make the fist row Bold
         */
        worksheet.getRow(1).font = {bold: true}

    } catch (error) {
        throw new Error(error);
    }
}

/**
 * Function to extract email and phone number from files and store in xlsx file
 * @param {*} file file name
 * @param {*} resolve promise resolve
 */
async function dataExtractorAndFileStorage(file, resolve){

    let dataBuffer, fileExt;
    try {
        fileExt = file.split('.');
    } catch (error) {
        throw new Error(error);
    }
    
    if( fileExt[fileExt.length - 1] == constFile.fileExtension){

        try {
            dataBuffer = fs.readFileSync(constFile.srcFolderPath+file);
        } catch (error) {
            throw new Error(error);
        }

        /**
         * PDF parser to extract data from pdf files.
         */
        pdf(dataBuffer).then(item => {
            
            try {
                
                /**
                 * Regex match for email and mobile number
                 */
                let email = item.text.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
                let phone = item.text.match(/(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|(\+\d{1,3}[- ]?)?\d{5}?[ ]?\d{5})/gi);

                let fileName = file.split('.')[0];
                let mailID, contact, altEmail, altPhone;

                if(email){
                    mailID = email[0];
                    altEmail = email.slice(1).join(', \n');
                }else{
                    mailID = null;
                    altEmail = null;
                }

                if(phone){
                    contact = phone[0];
                    altPhone = phone.slice(1).join(', \n');
                }else{
                    contact = null;
                    altPhone = null;
                }

                resolve(worksheet.addRow({name: fileName, id: mailID, altid: altEmail, num: contact, altnum: altPhone}));
            } catch (error) {
                throw new Error(error);
            }
        });
    }else{
        console.log(constFile.strInvalidFile + file);
        resolve();
    }
    
}

/**
 * Function to extract email and phone number from files and store in DB
 * @param {*} file file name
 * @param {*} resolve promise resolve
 */
async function dataExtractorAndDBStorage(file, resolve){
    
    let dataBuffer, fileExt;
    try {
        fileExt = file.split('.');
    } catch (error) {
        throw new Error(error);
    }
    
    if( fileExt[fileExt.length - 1] == constFile.fileExtension){

        try {
            dataBuffer = fs.readFileSync(constFile.srcFolderPath+file);
        } catch (error) {
            throw new Error(error);
        }

        /**
         * PDF parser to extract data from pdf files.
         */
        pdf(dataBuffer).then(item => {
            
            try {
                
                /**
                 * Regex match for email and mobile number
                 */
                let email = item.text.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
                let phone = item.text.match(/(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|(\+\d{1,3}[- ]?)?\d{5}?[ ]?\d{5})/gi);

                let fileName = file.split('.')[0];
                let mailID, contact, altEmail, altPhone;

                if(email){
                    mailID = email[0];
                    altEmail = email.slice(1);
                }else{
                    mailID = null;
                    altEmail = null;
                }

                if(phone){
                    contact = phone[0];
                    altPhone = phone.slice(1);
                }else{
                    contact = null;
                    altPhone = null;
                }
                return store(fileName, mailID, altEmail, contact, altPhone, resolve);
            } catch (error) {
                throw new Error(error);
            }
        });
    }else{
        console.log(constFile.strInvalidFile + file);
        resolve();
    }
}

/**
 * Function to save the workbook with all data
 */
function createWorkBook(){
    workbook1.xlsx.writeFile(constFile.destFilePath).then(function() {
        console.log(constFile.strFileWritten);
    }).catch((error)=>{
        throw new Error(error);
    })
}

function store(file_name, email, altEmail, contact, altContact, resolve){
    var body = {};
    body[constFile.strFileName] = file_name;
    body[constFile.strEmail] = email;
    body[constFile.strContact] = contact;
    body[constFile.strAltEmail] = altEmail;
    body[constFile.strAltContact] = altContact;
    var myData = new ResumeData(body);
    myData.save().then(item => {
        return resolve();
    }).catch(err => {
        console.log(err);
    });
}