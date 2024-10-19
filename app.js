const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const app = express();
const dotenv = require('dotenv');
const User = require('./model/user.model.js');
const db = require('./config/db');
dotenv.config();

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

app.post('/upload', upload.single('xlsxFile'), async (req, res) => {
    notsentemails = []; // Clear the array at the beginning of each request

    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }

    const senderEmail = req.body.senderEmail;
    const customMessage = req.body.customMessage;
    
    // Check if the user exists in the database
    let user = await User.findOne({ email: senderEmail });
    
    // If user does not exist, generate a password and store it
    if (!user) {
        const senderPassword = req.body.senderPassword;

        if (!isValidEmail(senderEmail)) {
            return res.status(400).send('Invalid sender email.');
        }

        // Save user to the database
        user = new User({
            email: senderEmail,
            password: senderPassword
        });
        await user.save();
    }

    // Use the stored password for authentication
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: senderEmail,
            pass: user.password // Use the saved hashed password
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
                        text: customMessage || 'That was easy!'
                    };
                    transporter.sendMail(mailOptions, function (error, info) {
                        if (error) {
                            notsentemails.push(value);
                            console.log('Error sending to:', value, error);
                        } else {
                            console.log('Email sent: ' + info.response);
                        }
                        resolve();
                    });
                } else {
                    resolve();
                }
            });
        });
    });

    // Wait for all email operations to finish
    Promise.all(emailPromises).then(() => {
        res.json(notsentemails); // Send the failed emails to the frontend
    });
});

app.get('/getPassword', async (req, res) => {
    const senderEmail = req.query.email;

    if (!isValidEmail(senderEmail)) {
        return res.status(400).send('Invalid email.');
    }

    // Find the user by email in the database
    const user = await User.findOne({ email: senderEmail });

    if (!user) {
        return res.status(404).send('User not found.');
    }

    // Send the password back to the frontend
    res.json({ password: user.password });
});


const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    db;
    console.log(`Server is running on port ${PORT}`);
});
