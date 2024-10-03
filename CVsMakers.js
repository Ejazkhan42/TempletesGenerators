const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");


const templateContent = fs.readFileSync(path.resolve(__dirname, "Template/Classic #3_table_V2.docx"), "binary");
const jsonDir = path.resolve(__dirname, "JSON");
const jsonFiles = fs.readdirSync(jsonDir).filter(file => file.endsWith('.json'));

function nullGetter(part) {
    if (part.raw) {
        return "";
    }
    if (!part.module && part.value) {
        return "";
    }
    return "";
}

const expressionParser = require("docxtemplater/expressions.js");
const { forEach } = require("jszip");
expressionParser.filters.upper = function (input) {
    if (!input) return input;
    return input.toUpperCase();
};

expressionParser.filters.clean = function (input) {
    if (!input) return "";

    if (!input) return [];
    const Fields={"Values":{}}
    const Lines={};
    const line=input
    .replace(/\*/g, '')
    .replace(/\t/g, '')
    .replace(/<p>/g, '')
    .replace(/<\/p>/g, '')
    .trim()
    .split('\n') 
    Fields["Values"]=line
    return Fields
}
expressionParser.filters.splitFields = function (input) {
    if (!input) return {};
    const fieldMappings = {
        'name': 'name',
        'email address': 'email',
        'phone contact': 'phone',
        'rfq': 'rfq',
        'clearance level': 'level'
    };

    const fields = {};

    const lines = input.split("\n");
    lines.forEach(line => {
        const [key, ...values] = line.split(":").map(str => str.trim());
        const value = values.join(':').trim();
        if (key.toLowerCase() in fieldMappings) {
            const fieldName = fieldMappings[key.toLowerCase()];
            fields[fieldName] = value;
        }
    });

    return fields;
};

jsonFiles.forEach(name => {
    const filePath = path.join(jsonDir, name);
    const cv = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    const zip = new PizZip(templateContent);
    const doc = new Docxtemplater(zip, {
        parser: expressionParser,
        nullGetter,
        paragraphLoop: true,
        linebreaks: true,
    });

    doc.setData(cv);
    doc.render();

    const buf = doc.getZip().generate({
        type: "nodebuffer",
        compression: "DEFLATE",
    });

    fs.writeFileSync(path.resolve(__dirname, `Files/${name.split('.')[0]}.docx`), buf);
    console.log(`Generated ${name.split('.')[0]}.docx`);
});

console.log("All files generated.");
