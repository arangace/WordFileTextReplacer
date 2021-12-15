var PizZip = require("pizzip");
var Docxtemplater = require("docxtemplater");

var fs = require("fs");
var path = require("path");

const prompt = require('prompt');

prompt.start();

//This function starts the user prompt and sends questions to the user
//The user then answers the questions of what information they want
//replaced and reads the users input and stores it for replacing later
const userPrompt = () => {
    //Define prompts to user
    const properties = [
        {
            name: 'original_file',
            description: 'Original file name (excluding .docx) if file name is example.docx, Example: example '
        },
        {

            name: 'receiver_name',
            description: 'Name of person addressed to',
            validator: /^[a-zA-Z\s\-]+$/,
            warning: 'Must contain only letters, spaces, or dashes',
            hidden: false
        },
        {
            name: 'event_name',
            description: 'Event title',
            validator: /^[a-zA-Z\s\-]+$/,
            warning: 'Event Title must be only letters, spaces, or dashes',
            hidden: false
        },
        {
            name: 'venue_name',
            description: 'Venue location',
            validator: /^[a-zA-Z\s\-]+$/,
            warning: 'Venue Title must be only letters, spaces, or dashes',
            hidden: false
        },
        {
            name: 'event_time',
            description: 'Event time in am or pm i.e. 2:00pm',
            validator: /^[0-23\a-zA-Z\s\:]+$/,
            warning: 'Event time in invalid format',
            hidden: false
        },
        {
            name: 'dress_code',
            description: 'Dress code',
            validator: /^[a-zA-Z\s\-]+$/,
            warning: 'Dress code must be only letters, spaces, or dashes',
            hidden: false
        },
        {
            name: 'senders_name',
            description: 'Sender name',
            validator: /^[a-zA-Z\s\-]+$/,
            warning: 'Sender must be only letters, spaces, or dashes',
            hidden: false
        }
    ];
    //sends message to console
    console.log("Please enter in the required details..")
    //receives console input and parses into replaceInformation for later handling
    prompt.get(
        properties,
        function (err, result) {
            let replaceInformation = []
            if (err) {
                return onErr(err);
            }
            properties.forEach((element) => {
                console.log(`${element.description} is: ${result[element.name]}`);
                replaceInformation[element.name] = result[element.name]
            })
            //Take the constructed data and replace the word file with the new data
            replace(replaceInformation);

        });

    function onErr(err) {
        console.log(err);
        return -1;
    }

}
//This function takes the new information to be replaced and replaces the
//old .docx file with the new information, creating a new file with the new information.
//Without tampering the original file.
const replace = (replaceInformation) => {
    // Load the docx file as binary content into the code for use
    try {
        var content = fs.readFileSync(
            path.resolve(__dirname, replaceInformation.original_file + ".docx"),
            "binary"
        );
    }
    catch (err) {
        console.log("\nERROR: Cannot read File or file does not exist\n");
        console.log('Exiting..')

        return -1;
    }

    var zip = new PizZip(content);
    //creates a new doc object, ready to take the new information
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    // render the document with the new replaced information i.e.
    // (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
    doc.render(
        replaceInformation
    );

    var buf = doc.getZip().generate({ type: "nodebuffer" });
    console.log(`Complete! \nFile created: ${replaceInformation[Object.keys(replaceInformation)[1]]}-${replaceInformation[Object.keys(replaceInformation)[2]]}.docx`)
    //Adds the new file in the specified directory with the new file name
    fs.writeFileSync(path.resolve(__dirname + "/wordFiles", ` ${replaceInformation[Object.keys(replaceInformation)[1]]}-${replaceInformation[Object.keys(replaceInformation)[2]]}.docx`), buf);
}
userPrompt();