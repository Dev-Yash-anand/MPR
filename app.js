const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const app = express();

var notsentemails = [];

function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, 'uploads/');
    },
    filename: function (req, file, cb) {
        cb(null, file.originalname);
    }
});

const upload = multer({ storage: storage });

app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

app.post('/upload', upload.single('xlsxFile'), (req, res) => {
    notsentemails = []; // Clear the array at the beginning of each request

    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    const senderEmail = req.body.senderEmail;
    const senderPassword = req.body.senderPassword;
    const customMessage = req.body.customMessage; // Capture the custom message

    if (!isValidEmail(senderEmail)) {
        return res.status(400).send('Invalid sender email.');
    }

    var transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: senderEmail,
            pass: senderPassword
        }
    });

    const results = [];
    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    // Extract emails from the spreadsheet
    data.forEach((row, rowIndex) => {
        const rowData = {};
        row.forEach((cell, columnIndex) => {
            rowData[`Column${columnIndex + 1}`] = cell;
        });
        results.push(rowData);
    });

    let emailPromises = results.map(item => {
        return new Promise((resolve) => {
            Object.values(item).forEach(value => {
                if (isValidEmail(value)) {
                    var mailOptions = {
                        from: senderEmail,
                        to: String(value),
                        subject: 'Sending Email using Node.js',
                        text: customMessage || 'That was easy!' // Use the custom message here
                    };
                    transporter.sendMail(mailOptions, function (error, info) {
                        if (error) {
                            notsentemails.push(value); // Add failed email to array
                            console.log('Error sending to:', value, error);
                        } else {
                            console.log('Email sent: ' + info.response);
                        }
                        resolve(); // Resolve after the email is processed
                    });
                } else {
                    resolve(); // Skip if it's not a valid email
                }
            });
        });
    });

    // Wait for all email operations to finish
    Promise.all(emailPromises).then(() => {
        res.json(notsentemails); // Send the failed emails to the frontend
    });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
