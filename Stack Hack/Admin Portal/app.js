const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Load Excel data
const workbook = xlsx.readFile('./credentials.xlsx');
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const credentials = xlsx.utils.sheet_to_json(worksheet);

// Helper function to verify credentials
function validateUser(username, password) {
    return credentials.some(user => user.username === username && user.password === password);
}

// Routes
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'views', 'login.html'));
});

app.post('/login', (req, res) => {
    const { username, password } = req.body;
    if (validateUser(username, password)) {
        res.redirect('/admin');
    } else {
        res.send('<h3>Invalid credentials. Please <a href="/">try again</a>.</h3>');
    }
});

app.get('/admin', (req, res) => {
    res.sendFile(path.join(__dirname, 'views', 'admin.html'));
});

app.listen(3000, () => {
    console.log('Server is running on http://localhost:3000');
});
