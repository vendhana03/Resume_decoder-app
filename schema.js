/**
 * Framework imports
 */
const mongoose = require('mongoose');

/**
 * Feature imports
 */
const constFile = require("./constant");

/**
 * Data Schema
 */
var DataSchema = new mongoose.Schema({
    file_name: {
        type: String,
        required:true
    },
    created_at: {
        type: Date,
        default: Date.now(),
        required: true
    },
    email: String,
    contact: String,
    alt_email: [ String ],
    alt_contact: [ String ]
});

module.exports = mongoose.model(constFile.strResumeData, DataSchema);