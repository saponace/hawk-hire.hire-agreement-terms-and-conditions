var EMAIL_SEND_AS_NAME = "Hire agreement automatic email sender";
var EMAIL_BODY = "Please find attached the agreement form";
var EMAIL_SUBJECT = "Agreement form";
var TABLE_PROPERTIES = {
    date: 'Date',
    customer: 'Customer',
    equipment: 'Equipment',
    location: 'Location',
    siteContact: 'Site Contact',
    agreementFormNo: 'Agreement Form #'
};

/**
 * Display the email selection popup
 * This popup then calls the sendAgreementForm function with the form data as a parameter
 */
function showEmailSelectionPopup() {
    var htmlTemplate = HtmlService.createTemplateFromFile('email-selection-popup');
    var htmlOutput = htmlTemplate.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle('Save as PDF and send Hire Agreement')
        .setWidth(230)
        .setHeight(120);
    DocumentApp.getUi().showDialog(htmlOutput);
}

/**
 * Send the agreement form
 * @param data the data output from the HTML form
 */
function sendAgreementForm(data) {
    var recipientEmailAddress = data.emailAddress;
    if(recipientEmailAddress) {
        console.log("Exporting the document to PDF ...")
        var fileName = getExportPdfName();
        var exportFile = createPdfToExport(fileName);
        console.log("Export done")

        console.log("Sending the email")
        sendEmail(recipientEmailAddress, exportFile);
        console.log("Email sent")

        console.log("Incrementing the agreement form number")
        incrementAgreementFormNo();
        console.log("Agreement form number incremented")

        console.log("Wiping data table")
        wipeDataTable();
        console.log("Data table wiped")
    }
}

//2. save as PDF in folder with the name "HIRE AGREEMENT FORM # XXX (last row of header table) - CUSTOMER (2nd row of header table)
/**
 * Convert the doc to PDF and save that PDF in the export folder
 * @param fileName Name to give to the pdf export file
 * @return the created file
 */
function createPdfToExport(fileName) {
    var fileName = fileName;
    var docBlob = DocumentApp.getActiveDocument().getAs('application/pdf');
    docBlob.setName(fileName+ ".pdf");

    var outputFolder = DriveApp.getFolderById(EXPORT_FOLDER_ID);
    return outputFolder.createFile(docBlob);
}

/**
 * Send an email with attachment
 * @param emailAddress Recipient to send the email to
 * @param attachment File to attach to the email
 */
function sendEmail(emailAddress, attachment) {
    var recipient = emailAddress;
    var subject = EMAIL_SUBJECT;
    var body = EMAIL_BODY;
    var emailOptions = {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: EMAIL_SEND_AS_NAME
    };
    MailApp.sendEmail(recipient, subject, body, emailOptions);
}

/**
 * Get the paragraph representing the table cell that corresponds to the given rproperty (looks for the parameter string
 * in the first column, returns the second column cell
 * @param propName Content of the first column cell
 * @returns The paragraph object if it exists, null otherwise
 */
function getPropParagraph(propName) {
    var paragraphs = DocumentApp.getActiveDocument().getBody().getParagraphs();
    for (var i = 0; i < paragraphs.length; i++) {
        var p = paragraphs[i];
        if(p.getText() === propName){
            if(i < paragraphs.length - 1)
                return paragraphs[i+1];
            else
                return null;
        }
    }
    return null;
}

/**
 * Bulids the output PDF filename
 * @return The name
 */
function getExportPdfName() {
    var agreementNo = getPropParagraph(TABLE_PROPERTIES.agreementFormNo).getText();
    var customerName = getPropParagraph(TABLE_PROPERTIES.customer).getText();
    return "Hire agreement form #" + agreementNo + " - " + customerName;
}

/**
 * increment the agreement form number in the data table
 */
function incrementAgreementFormNo() {
    var agreementNoParagraph = getPropParagraph(TABLE_PROPERTIES.agreementFormNo);
    var agreementNo = parseInt(agreementNoParagraph.getText());
    var incrementedAgreementNo = agreementNo + 1;
    var incrementedAgreementNoString = addZeroPadding(incrementedAgreementNo, 4);
    agreementNoParagraph.setText(incrementedAgreementNoString);
}

/**
 * Add zeroes padding to a string
 * @param inputStr The string to pad
 * @param width The total number of chars to reach
 * @return {string|*} The padded string
 */
function addZeroPadding(inputStr, width) {
    inputStr = inputStr + '';
    return inputStr.length >= width ? inputStr : new Array(width - inputStr.length + 1).join('0') + inputStr;
}

/**
 * Wipe values cells in the data table
 */
function wipeDataTable() {
    getPropParagraph(TABLE_PROPERTIES.date).setText(' ');
    getPropParagraph(TABLE_PROPERTIES.customer).setText(' ');
    getPropParagraph(TABLE_PROPERTIES.equipment).setText(' ');
    getPropParagraph(TABLE_PROPERTIES.location).setText(' ');
    getPropParagraph(TABLE_PROPERTIES.siteContact).setText(' ');
}