
const express = require('express');


const app = express();
app.use(express.json());

const multer = require('multer');
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });


app.post('/generate-docx', upload.fields([{ name: 'excel' }, { name: 'docx' }]), async (req, res) => {

// npm use
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');

    const { foldername } = req.body;
    console.log(req.body);

    res.status(200).send('Files processed successfully');


    const rootFolderPath = path.join('hdfc/', foldername);
    console.log("Folder path:", rootFolderPath);

    if (!fs.existsSync(rootFolderPath)) {
        fs.mkdirSync(rootFolderPath, { recursive: true });
    }

    try {
        const excelFile = req.files['excel'][0];
        const docxFile = req.files['docx'][0];

        const workbook = XLSX.read(excelFile.buffer);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const excelData = XLSX.utils.sheet_to_json(sheet);

        for (let index = 0; index < excelData.length; index++) {
            const row = excelData[index];
            const docxBuffer = docxFile.buffer; // Get the buffer for the DOCX file

            const doc = new Docxtemplater(new PizZip(docxBuffer));



            function excelDateToJSDate(serial) {
                var excelStartDate = new Date(Date.UTC(1899, 11, 30)); // Set the start date
                var actualDate = new Date(excelStartDate.getTime() + serial * 86400000); // Convert serial to milliseconds
                return actualDate.toISOString().substring(0, 10); // Return date in YYYY-MM-DD format
            }

            const formattedNoticeDate = excelDateToJSDate(row['notice date']);
            const formattedNPADate = excelDateToJSDate(row['NPA DATE']);
            const formattedCLAIM_AMOUNT_AS_ONDate = excelDateToJSDate(row['CLAIM AMOUNT AS ON']);

            const customerFatherHusbandName = row['CUSTOMER FH NAME'] || 'Not Available';

            console.log({
                "ed_number": row['ed number'],
                "Notice_date": formattedNoticeDate,
                "lot": row['lot'],
                "file_no": row['file no.'],
                "CUSTOMER_NAME": row['CUSTOMER NAME'],
                "CUSTOMER_NAME2": row['CUSTOMER NAME2'],
                //"CUSTOMER_FATHER_HUSBAND_NAME": row['CUSTOMER FATHER OR HUSBAND NAME'],
                "CUSTOMER_FATHER_HUSBAND_NAME": customerFatherHusbandName,
                "CUSTOMER_ADDRESS1_WITH_PIN_CODE": row['CUSTOMER_ADDRESS1_WITH_PIN_CODE'],
                "PRODUCT": row['PRODUCT'],
                "LOAN_ACCOUNT_NO": row['LOAN ACCOUNT NO'],
                "DISBURSEMENT_AMOUNT": row['DISBURSEMENT AMOUNT'],
                "NPA_DATE": formattedNPADate,
                "CLAIM_AMOUNT": row['CLAIM AMOUNT'],
                "amount_in_words": row['amount in words'],
                "CLAIM_AMOUNT_AS_ON": formattedCLAIM_AMOUNT_AS_ONDate,
                "BU_BRANCH_NAME": row['BU/ BRANCH NAME'],
                "REGIONTERRITORY": row['REGION/TERRITORY'],
                "TEST": row['test'],
                "COLLECTION_OFFICER_NAME": row['COLLECTION OFFICER NAME'],
                "COLLECTION_OFFICER_MOBILE_NUMBER": row['COLLECTION OFFICER MOBILE NUMBER']
            });

            doc.setData({
                "ed_number": row['ed number'],
                "Notice_date": formattedNoticeDate,
                "lot": row['lot'],
                "file_no": row['file no.'],
                "CUSTOMER_NAME": row['CUSTOMER NAME'],
                "CUSTOMER_NAME2": row['CUSTOMER NAME2'],
               
                "CUSTOMER FATHER OR HUSBAND NAME": customerFatherHusbandName,
                "CUSTOMER_ADDRESS1_WITH_PIN_CODE": row['CUSTOMER_ADDRESS1_WITH_PIN_CODE'],
                "PRODUCT": row['PRODUCT'],
                "LOAN_ACCOUNT_NO": row['LOAN ACCOUNT NO'],
                "DISBURSEMENT_AMOUNT": row['DISBURSEMENT AMOUNT'],
                "NPA_DATE": formattedNPADate,
                "CLAIM_AMOUNT": row['CLAIM AMOUNT'],
                "amount_in_words": row['amount in words'],
                "CLAIM_AMOUNT_AS_ON": formattedCLAIM_AMOUNT_AS_ONDate,
                "TEST": row['test'],
                "BU_BRANCH_NAME": row['BU/ BRANCH NAME'],
                "REGIONTERRITORY": row['REGION/TERRITORY'],
                "COLLECTION_OFFICER_NAME": row['COLLECTION OFFICER NAME'],
                "COLLECTION_OFFICER_MOBILE_NUMBER": row['COLLECTION OFFICER MOBILE NUMBER']
            });
console.log(customerFatherHusbandName)
            doc.render();

            const docxUpdatedContent = doc.getZip().generate({ type: 'nodebuffer' });
            const pdfDoc = await convertDocxToPdf(docxUpdatedContent); // Pass docxUpdatedContent to conversion function
            const outputPath = path.resolve(rootFolderPath, `Updated_Doc_${index + 1}.pdf`);

            fs.writeFileSync(outputPath, pdfDoc);
        }

    } catch (error) {
        console.error('Error processing files:', error);
        res.status(500).send('Error processing files');
    }
});

async function convertDocxToPdf(docxBuffer) {
    const mammoth = require("mammoth");
    const htmlToPdf = require("html-pdf");

    // Convert the DOCX content to HTML
    let { value } = await mammoth.convertToHtml({ buffer: docxBuffer });

    // Removing specific unwanted lines
    const unwantedPatterns = [
        "KONCEPT LAW Ambika Mehra",
        "ASSOCIATES \\(Advocate\\)",
        "B\\.Sc\\. LL\\.B"
    ];

    unwantedPatterns.forEach(pattern => {
        // This regex accounts for potential HTML tags and spaces around the words
        let regex = new RegExp(pattern.split(" ").join("\\s*(?:<[^>]+>\\s*)?"), 'gi');
        value = value.replace(regex, '');
    });

    // Correct undefined entries


    // Find the position to insert the signature image
    const signatureHtml = '<p><img src="https://raw.githubusercontent.com/adityagithubraj/pinterest_clone/main/photo/WhatsApp%20Image%202024-04-18%20at%2016.41.00_fdf16bc1.jpg" alt="Signature" style="width: 100px; height: auto;"></p>';
    const closingText = "Yours faithfully,";
    const closingPosition = value.indexOf(closingText);

    // Insert the signature after "Yours faithfully,"
    let enhancedHtml;
    if (closingPosition !== -1) {
        enhancedHtml = value.slice(0, closingPosition + closingText.length) + signatureHtml + value.slice(closingPosition + closingText.length);
    } else {
        enhancedHtml = value + signatureHtml;  // Fallback if the specific text isn't found
    }

    // Center "Loan Recall Notice" and the date
    enhancedHtml = enhancedHtml.replace(/Loan Recall Notice/g, '<center>Loan Recall Notice</center>');
    //enhancedHtml = enhancedHtml.replace(/Date :\s+([0-9]+)/g, (match, p1) => `<center>Date : ${p1}</center>`);

    // Enhanced styles and final HTML setup
    const styledHtml = `
    <style>
    body {
        font-family: 'Times New Roman', serif;
        font-size: 8pt;
        color: #333;
        margin: 50px;
        line-height: 1.3;
    }
    h1 {
        color: #000;
        font-size: 18pt;
        text-align: center;
        font-weight: bold;
    }
    table.header-table {
        width: 100%;
        margin-bottom: 20px; /* Space between header and content */
        border-collapse: collapse;
        border: none;
    }
    table.header-table td {
        border: none; /* Remove borders specifically from header cells */
    }
    .header, .header2 {
        font-size: 20pt;
        padding: 0;
        margin: 0;
        display: flex;
        flex-direction: column;
        font-weight: bold;
    }
    .content {
        text-align: justify;
    }
    .ref-details, .address {
        font-size: 10pt;
        margin-left: 0;
        margin-bottom: 15px;
        font-weight: bold;
    }
    .footer {
        font-size: 8pt;
        text-align: center;
        position: fixed;
        bottom: 0;
        width: 100%;
    }
    table, td, th {
        border: 1px solid red;
        font-size: 8pt;
        border-collapse: collapse;
        padding:0
    }
    th{
        border: 1px solid blue;
        font-size: 12pt;
        border-collapse: collapse;
        padding:0
    }
    td{
        border: 1px solid black;
        font-size: 8pt;
        border-collapse: collapse;
        padding:10px
    }
   
    td th{
        margin-top:-100px;
        border: 1px solid red;
    }
    center {
        display: block;
        margin-top: 0;
        margin-bottom: 0;
        text-align: center;
    }

    </style>
    <table class="header-table">
        <tr>
            <td class="header">
                KONCEPT LAW <br> ASSOCIATES
            </td>
            <td class="header2" style="text-align: right;">
             Ambika     Mehra <br> 
                <div style="font-size: 10pt;">(Advocate)</div> 
                <p style="font-size: 8pt;">B.Sc. LL.B</p>
            </td>
        </tr>
    </table>
    ${enhancedHtml}
`;

    const options = { format: "A4", border: { top: '0mm', bottom: '2mm', left: '5mm', right: '5mm' } };
    return new Promise((resolve, reject) => {
        htmlToPdf.create(styledHtml, options).toBuffer((err, buffer) => {
            if (err) {
                reject(err);
            } else {
                resolve(buffer);
            }
        });
    });
}



app.listen(5000, () => {
    console.log(`Server is listening at http://localhost:5000`);
});