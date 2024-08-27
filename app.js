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

var transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: 'anshumanpandey182005@gmail.com',
        pass: 'vogn kquw wyir klag'
    }
});

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
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }
    const results = [];
    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
    data.forEach((row, rowIndex) => {
        const rowData = {};
        row.forEach((cell, columnIndex) => {
            rowData[`Column${columnIndex + 1}`] = cell;
        });
        results.push(rowData);
    });
    results.map(item => {
        Object.values(item).forEach(value => {
            if (isValidEmail(value)) {
                var mailOptions = {
                    from: 'anshumanpandey182005@gmail.com',
                    to: String(value),
                    subject: 'Sending Email using Node.js',
                    text: 'That was easy!'
                };
                transporter.sendMail(mailOptions, function(error, info){
                    if (error) {
                        notsentemails.push(value);
                        console.log(error);
                    } else {
                        console.log('Email sent: ' + info.response);
                    }
                });
            }
        });
    });
    res.json(notsentemails);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});