# probable-waddle

README for XML to DOCX Batch Script

///Overview///
This script parses XML files, extracts relevant information, and populates DOCX templates with the data. It generates personalized billing letters in DOCX format for each entry found in the XML file.

///Dependencies///
To run this script, you'll need to install the following dependencies:

xmldom - for parsing XML data.
pizzip - for handling ZIP files, which is necessary to manipulate DOCX templates.
docxtemplater - for generating DOCX documents based on templates.

///Installation///
Before running the script, make sure you have Node.js installed. Then, run the following commands to install the required packages:

npm install xmldom pizzip docxtemplater

///Project Structure///
assets/CurrentBillData.xml – The sample XML file containing the billing data.
assets/CurrentBillTemplate.docx – The DOCX template to be filled with data.
TOTRANSLATE/ – The folder where the generated DOCX files will be saved.

///How to Run the Script///
Place the XML data file (CurrentBillData.xml) and DOCX template (CurrentBillTemplate.docx) inside the assets folder.
Run the script using Node.js:

node XML-extractor_v2.js

The script will process the XML data, populate the DOCX template with the extracted data, and save the generated documents in the TOTRANSLATE folder. Each file will be named according to the documentId extracted from the XML.

///Output///
The generated DOCX files are saved in the TOTRANSLATE/ folder with filenames in the format UpdatedBill\_<documentId>.docx.

///Notes///
Ensure that the structure of the XML file matches the expected format in the script.
The folder TOTRANSLATE/ will be created automatically if it doesn't exist.
Script can be expanded to automatically detect doc type and retrieve corresponding templates.
Script can be encapsulated into an executable (i.e., .exe for Windows, .deb for Linus, or .pkg for MacOS).
