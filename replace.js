var PizZip = require("pizzip");
var Docxtemplater = require("docxtemplater");

var fs = require("fs");
var path = require("path");

const prompt = require('prompt');

prompt.start();

const userPrompt = () => {
    const replaceInformation = []
    const properties = [
        {
            name: 'company_name',
            validator: /^[a-zA-Z\s\-]+$/,
            warning: 'Username must be only letters, spaces, or dashes',
            hidden: false
        },
        {
            name: 'job_title',
            validator: /^[a-zA-Z\s\-]+$/,
            warning: 'Username must be only letters, spaces, or dashes',
            hidden: false
        }
    ];
    prompt.get(
        properties,
        function (err, result) {
            if (err) {
                return onErr(err);
            }
            console.log('Company data successfully recieved');
            console.log('  Company name: ' + result.company_name);
            console.log('  Job Title: ' + result.job_title);
            replaceInformation.company_name = result.company_name
            replaceInformation.job_title = result.job_title
            console.log(replaceInformation)
        });

    function onErr(err) {
        console.log(err);
        return 1;
    }

    replace(replaceInformation);
}
const replace = (replaceInformation) => {

    // Load the docx file as binary content
    var content = fs.readFileSync(
        path.resolve(__dirname, "simple.docx"),
        "binary"
    );
    var zip = new PizZip(content);

    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });


    // render the document
    // (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
    doc.render({
        company_name: replaceInformation.company_name,
        job_title: replaceInformation.job_title,
    });

    var buf = doc.getZip().generate({ type: "nodebuffer" });

    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    fs.writeFileSync(path.resolve(__dirname, `${company_name, job_title}-cover-letter.docx`), buf);
}
userPrompt();