const fs = require('fs');
const { DOMParser } = require('xmldom');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const path = require('path');

// Load the XML data
const xmlData = fs.readFileSync('assets/CurrentBillData.xml', 'utf-8');
const doc = new DOMParser().parseFromString(xmlData, 'text/xml');

// Define the path to the assets and OUT folder
const assetsFolder = path.join(__dirname, 'assets');
const outFolder = path.join(__dirname, 'TOTRANSLATE');

// Create the OUT folder if it doesn't exist
if (!fs.existsSync(outFolder)) {
    fs.mkdirSync(outFolder);
}

// Get all <medicare_batch_bills_letters> nodes
const medicareLetters = doc.getElementsByTagName('medicare_batch_bills_letters');

// Log the number of members
console.log(`Number of members in the file: ${medicareLetters.length}`);

// Extract data for each <medicare_batch_bills_letters> section
let extractedData = [];
for (let i = 0; i < medicareLetters.length; i++) {
    const letterData = extractData(medicareLetters[i]);
    extractedData.push(letterData);
}

// Sort the array by documentId or any other attribute
extractedData.sort((a, b) => a['documentId'].localeCompare(b['documentId']));

// Save the result to a file
fs.writeFileSync('sortedMedicareBatchBills.json', JSON.stringify(extractedData, null, 2));

// Load the DOCX template
const content = fs.readFileSync(path.join(assetsFolder, 'CurrentBillTemplate.docx'), 'binary');
const zip = new PizZip(content);

// Now, for each extracted letter, generate a DOCX file with the updated first name
extractedData.forEach(letterData => {
    // [TEST] Log the data structure to verify keys
    //console.log(letterData);

    // Correct the path to first_name based on the actual structure
    const documentId = letterData.documentId;
    const firstName = letterData.batch_letter_gen['batch_letter_gen.subscriber_name']['batch_letter_gen.subscriber_name.first_name'];
    const lastName = letterData.batch_letter_gen['batch_letter_gen.subscriber_name']['batch_letter_gen.subscriber_name.last_name'];
    const groupNumber = letterData.batch_letter_gen['batch_letter_gen.group_number'];
    const accountNumber = letterData.batch_letter_gen['batch_letter_gen.collective_id'];
    const memberNumber = letterData.batch_letter_gen['batch_letter_gen.subscriber_id'];
    const invoiceNumber = letterData.batch_letter_gen['batch_letter_gen.invoice_number'];
    const sendAddress1 = letterData.batch_letter_gen['batch_letter_gen.send_address']['batch_letter_gen.send_address.address_1'];
    const sendAddress2 = letterData.batch_letter_gen['batch_letter_gen.send_address']['batch_letter_gen.send_address.address_2'];
    const sendCity = letterData.batch_letter_gen['batch_letter_gen.send_address']['batch_letter_gen.send_address.city'];
    const sendState = letterData.batch_letter_gen['batch_letter_gen.send_address']['batch_letter_gen.send_address.state'];
    const sendZip = letterData.batch_letter_gen['batch_letter_gen.send_address']['batch_letter_gen.send_address.zip'];
    const returnAddress1 = letterData.batch_letter_gen['batch_letter_gen.return_address']['batch_letter_gen.return_address.address_1'];
    const returnAddress2 = letterData.batch_letter_gen['batch_letter_gen.return_address']['batch_letter_gen.return_address.address_2'];
    const returnAddress3 = letterData.batch_letter_gen['batch_letter_gen.return_address']['batch_letter_gen.return_address.address_3'];
    const returnCity = letterData.batch_letter_gen['batch_letter_gen.return_address']['batch_letter_gen.return_address.city'];
    const returnState = letterData.batch_letter_gen['batch_letter_gen.return_address']['batch_letter_gen.return_address.state'];
    const returnZip = letterData.batch_letter_gen['batch_letter_gen.return_address']['batch_letter_gen.return_address.zip'];

    const memberFullName = letterData.batch_letter_gen['batch_letter_gen.covered_members']['batch_letter_gen.covered_members.member_full_name'];

    // Implementation of formatSystemDate
    const systemDate = formatSystemDate(letterData.batch_letter_gen['batch_letter_gen.system_date']);

    // Implementation of formatToMMDDYYYY
    const billingFirstDay = formatToMMDDYYYY(letterData.batch_letter_gen['batch_letter_gen.first_day_of_billing_cycle']);
    const billingLastDay = formatToMMDDYYYY(letterData.batch_letter_gen['batch_letter_gen.last_day_of_billing_cycle']);
    const billDueDate = formatToMMDDYYYY(letterData.batch_letter_gen['batch_letter_gen.bill_due_date']);
    const billDate = formatToMMDDYYYY(letterData.batch_letter_gen['batch_letter_gen.bill_date']);
    const billStartDate = formatToMMDDYYYY(letterData.batch_letter_gen['batch_letter_gen.biling_period_start_date']);
    const billEndDate = formatToMMDDYYYY(letterData.batch_letter_gen['batch_letter_gen.biling_period_end_date']);
    const coverageEffectiveDate = formatToMMDDYYYY(letterData.batch_letter_gen['batch_letter_gen.covered_members']['batch_letter_gen.covered_members.coverage_effective_date']);

    // Implementation of formatToDollar
    const previousAmountDue = formatToDollar(letterData.batch_letter_gen['batch_letter_gen.previous_amount_due']);
    const payments = formatToDollar(letterData.batch_letter_gen['batch_letter_gen.payments']);
    const currentCharges = formatToDollar(letterData.batch_letter_gen['batch_letter_gen.current_charges']);
    const totalDue = formatToDollar(letterData.batch_letter_gen['batch_letter_gen.amount_owed']);

    // Clone the template for each letter
    const docx = new Docxtemplater(zip.clone());

    // Set the data for the placeholder in the template
    docx.setData({
        firstName: firstName,
        lastName: lastName,
        groupNumber: groupNumber,
        accountNumber: accountNumber,
        memberNumber: memberNumber,
        invoiceNumber: invoiceNumber,

        systemDate: systemDate,
        billingFirstDay: billingFirstDay,
        billingLastDay: billingLastDay,
        billDueDate: billDueDate,
        billDate: billDate,
        bDate: billStartDate,
        bDate: billEndDate,

        prev: previousAmountDue,
        pay: payments,
        curr: currentCharges,
        due: totalDue,

        memberFullName: memberFullName,
        coverageEffectiveDate: coverageEffectiveDate,

        sendAddress1: sendAddress1,
        sendAddress2: sendAddress2,
        sendCity: sendCity,
        sendState: sendState,
        sendZip: sendZip,
        returnAddress1: returnAddress1,
        returnAddress2: returnAddress2,
        returnAddress3: returnAddress3,
        returnCity: returnCity,
        returnState: returnState,
        returnZip: returnZip
    });

    try {
        // Render the document
        docx.render();
    } catch (error) {
        console.error('Error rendering the DOCX:', error);
        return;
    }

    // Save the updated document with the correct name in the OUT folder
    // const outputPath = path.join(outFolder, `UpdatedBill_${documentId}.docx`);
    // const buf = docx.getZip().generate({ type: 'nodebuffer' });
    // fs.writeFileSync(outputPath, buf);
});

console.log('Documents successfully created in the TOTRANSLATE folder.');

//////////////////////////////////////////////////////////////////////////////////////
//////////////////////////// Data manipulation functions /////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////

// Function to recursively extract data from an XML node
function extractData(node, parentPath = '') {
    let data = {};

    // Iterate over child nodes
    for (let i = 0; i < node.childNodes.length; i++) {
        const child = node.childNodes[i];

        // Skip non-element nodes (like text nodes or comments)
        if (child.nodeType !== 1) continue;

        // Build the path using parent tag
        const nodePath = parentPath ? `${parentPath}.${child.nodeName}` : child.nodeName;

        if (child.childNodes.length > 1) {
            // If the node has children, recurse into them
            data[nodePath] = extractData(child, nodePath);
        } else {
            // If it's a leaf node, get its text content
            data[nodePath] = child.textContent.trim();
        }
    }

    return data;
}

// Function to format dates from YYYY-MM-DD to Month DD, YYYY
function formatSystemDate(dateString) {
    const monthNames = [
        "January", "February", "March", "April", "May", "June", "July",
        "August", "September", "October", "November", "December"
    ];

    // Split the date string manually
    const [year, month, day] = dateString.split('-');

    // Convert the month from 0-indexed (JavaScript Date behavior) to 1-indexed
    const formattedDate = `${monthNames[parseInt(month) - 1]} ${parseInt(day)}, ${year}`;

    return formattedDate;
}

// Function to format dates from YYYY-MM-DD to MM/DD/YYYY
function formatToMMDDYYYY(dateString) {
    // Split the date string manually
    const [year, month, day] = dateString.split('-');

    // Return in MM/DD/YYYY format
    return `${month}/${day}/${year}`;
}

function formatToDollar(value) {
    // Ensure the value is correctly parsed as a float, even if it's negative
    const numberValue = parseFloat(value);

    // Check if the value is negative and format accordingly
    if (numberValue < 0) {
        return `-$${Math.abs(numberValue).toFixed(2)}`; // Keeps 2 decimal places
    } else {
        return `$${numberValue.toFixed(2)}`; // Keeps 2 decimal places for positive values
    }
}