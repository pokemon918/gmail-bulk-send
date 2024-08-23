require('dotenv').config(); // Load environment variables from .env file
const { google } = require('googleapis');
const fs = require('fs');
const mammoth = require('mammoth');
const xlsx = require('xlsx');

// Function to read the subject from a text file
const getSubject = async (filePath) => {
  try {
    // Read the file synchronously
    const subject = fs.readFileSync(filePath, 'utf8');
    return subject;
  } catch (err) {
    console.error('Error reading the file:', err);
  }
};

// Function to get HTML content from a .docx file
const getHtmlContent = async (filePath) => {
  try {
    // Read the .docx file and convert it to HTML
    const { value: htmlContent } = await mammoth.convertToHtml({ path: filePath });
    return htmlContent; // Return the HTML content
  } catch (error) {
    console.error('Error reading the .docx file:', error);
  }
};

// Function to read data from an Excel file
const readExcelData = (filePath) => {
  try {
    const workbook = xlsx.readFile(filePath);
    const sheetName = 'Sheet1';
    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
      console.error(`Sheet "${sheetName}" does not exist.`);
      return [];
    }

    const data = [];
    let rowIndex = 1; // Start from row 1 (A1)
    while (true) {
      const cellAddress = `A${rowIndex}`;
      const cell = sheet[cellAddress];
      if (!cell) {
        break;
      }
      data.push(cell.v);
      rowIndex++;
    }
    return data;
  } catch (error) {
    console.error('Error reading the Excel file:', error);
    return [];
  }
};

// Function to send emails
const sendEmail = async () => {
  try {
    const arrayBody = readExcelData("assets/Email.xlsx");
    const oAuth2Client = new google.auth.OAuth2(
      process.env.CLIENT_ID,
      process.env.CLIENT_SECRET,
      process.env.REDIRECT_URI
    );
    await oAuth2Client.setCredentials({ refresh_token: process.env.REFRESH_TOKEN });
    const gmail = google.gmail({ version: 'v1', auth: oAuth2Client });
    const BATCH_LIMIT = 500; // Gmail API recommends sending no more than 500 emails at once
    const delayBetweenBatches = 60000; // 1 minute delay between batches

    for (let i = 0; i < arrayBody.length; i += BATCH_LIMIT) {
      const batch = arrayBody.slice(i, i + BATCH_LIMIT);
      await Promise.all(batch.map(async (recipient) => {
        try {
          const subject = await getSubject("assets/Subject.txt"); // Await the subject
          const utf8Subject = `=?utf-8?B?${Buffer.from(subject).toString('base64')}?=`;
          const htmlContent = await getHtmlContent("assets/Body.docx"); // Await the HTML content

          const messageParts = [
            'From: ' + process.env.FROM_EMAIL,
            'To: ' + recipient,
            'Content-Type: text/html; charset=utf-8',
            'MIME-Version: 1.0',
            `Subject: ${utf8Subject}`,
            '',
            htmlContent // Use the HTML content
          ];
          const message = messageParts.join('\n');
          const encodedMessage = Buffer.from(message)
            .toString('base64')
            .replace(/\+/g, '-')
            .replace(/\//g, '_')
            .replace(/=+$/, '');

          const response = await gmail.users.messages.send({
            userId: 'me',
            requestBody: { raw: encodedMessage },
          });
          console.log(`Email sent to: ${recipient}, Response: ${response.data}`);
        } catch (err) {
          console.error(`Failed to send email to ${recipient}:`, err);
        }
      }));

      // Delay before sending the next batch
      if (i + BATCH_LIMIT < arrayBody.length) {
        console.log(`Waiting for ${delayBetweenBatches / 1000} seconds before sending the next batch...`);
        await new Promise(resolve => setTimeout(resolve, delayBetweenBatches));
      }
    }
    console.log('All emails have been processed successfully.');
  } catch (error) {
    console.error('Error in sendEmail:', error);
  }
};

// Execute the sendEmail function
sendEmail();