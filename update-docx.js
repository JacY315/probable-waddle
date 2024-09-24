const fs = require('fs');
const { DOMParser } = require('xmldom');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const path = require('path');

// Define the path to the assets and OUT folder
const assetsFolder = path.join(__dirname, 'assets');
const outFolder = path.join(__dirname, 'OUT');

// Create the OUT folder if it doesn't exist
if (!fs.existsSync(outFolder)) {
    fs.mkdirSync(outFolder);
}

// Load the XML data
const xmlData = fs.readFileSync(path.join(assetsFolder, 'CurrentBillData.xml'), 'utf-8');

// Parse the XML
const doc = new DOMParser().parseFromString(xmlData, 'text/xml');

// Get all <subscriber_name> nodes
const subscribers = doc.getElementsByTagName('subscriber_name');

// Load the DOCX template
const content = fs.readFileSync(path.join(assetsFolder, 'CurrentBillTemplate.docx'), 'binary');
const zip = new PizZip(content);
const docxTemplate = new Docxtemplater(zip);

// Loop over subscribers to replace the placeholder for each subscriber
Array.from(subscribers).forEach((subscriber, index) => {
    const firstName = subscriber.getElementsByTagName('first_name')[0].textContent;

    // Clone the template for each subscriber
    const docx = new Docxtemplater(zip.clone());

    // Update the template for each first name
    docx.setData({
        firstName: firstName
    });

    try {
        // Render the document
        docx.render();
    } catch (error) {
        console.error('Error rendering the DOCX:', error);
    }

    // Save the updated document in the OUT folder
    const outputPath = path.join(outFolder, `UpdatedBill_${index}_${firstName}.docx`);
    const buf = docx.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(outputPath, buf);
});

console.log('Documents updated successfully in the OUT folder.');
