const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const fs = require("fs");
const path = require("path");

// Load the docx file as binary content
const content = fs.readFileSync(
    path.resolve(__dirname, "Demo_Template.docx"),
    "binary"
);


// const content = fs.readFileSync(
//     path.resolve(__dirname, "input.docx"),
//     "binary"
// );

const zip = new PizZip(content);


const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
});

// Render the document (Replace {first_name} by John, {last_name} by Doe, ...)
doc.render({
    userSex: "M",
    userAge: "18",
    userName: "John Doe",
    companyName: "ABC Comapny",
    userNational: "Hong Kong",

    textSmall: "Small",
    textMiddel: "Middel",
    textLarge: "Large",
    isCheck: true
});

const buf = doc.getZip().generate({
    type: "nodebuffer",
    // compression: DEFLATE adds a compression step.
    // For a 50MB output document, expect 500ms additional CPU time
    compression: "DEFLATE",
});

// buf is a nodejs Buffer, you can either write it to a
// file or res.send it with express for example.
fs.writeFileSync(path.resolve(__dirname, `output_${Date.now()}.docx`), buf);
console.log('End');