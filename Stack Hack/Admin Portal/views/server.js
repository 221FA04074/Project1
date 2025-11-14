// server.js
const express = require('express');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const cors = require("cors");
const app = express();

app.use(bodyParser.json());
app.use(cors());
// Load the Excel data
const workbook = xlsx.readFile('credentials.xlsx'); 
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const excelData = xlsx.utils.sheet_to_json(worksheet);
console.log(excelData);
app.post('/validate-login', (req, res) => {
    const { username, password } = req.body;
    const isValid = excelData.some(record => record.Username == username && record.Password == password);

    if (isValid) {
        res.json({ success: true});
    } else {
        res.json({ success: false});
    }
});

app.listen(3030, () => {
    console.log('Server is running on port 3030');
});
