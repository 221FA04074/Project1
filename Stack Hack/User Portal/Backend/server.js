// server.js
const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const cors = require("cors");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = 3000;

// Paths for separate files
const registrationFilePath = path.join(__dirname, "registration.xlsx");
const dataFilePath = path.join(__dirname, "data.xlsx");
const contactFilePath = path.join(__dirname, "contact.xlsx");
const extraServicesFilePath = path.join(__dirname, "extra_services.xlsx");

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Helper function to get or create a workbook
const getOrCreateWorkbook = (filePath, sheetName, headers) => {
    let workbook;
    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
    } else {
        workbook = xlsx.utils.book_new();
        const worksheet = xlsx.utils.aoa_to_sheet([headers]);
        xlsx.utils.book_append_sheet(workbook, worksheet, sheetName);
        xlsx.writeFile(workbook, filePath);
    }
    return workbook;
};

// Route to handle registration data submission
app.post("/register", (req, res) => {
    const { name, username, password, email, phone } = req.body;
    const workbook = getOrCreateWorkbook(registrationFilePath, "Registrations", ['Name', 'Username', 'Password', 'Email', 'Phone']);
    const worksheet = workbook.Sheets["Registrations"];

    const newRow = [name, username, password, email, phone];
    const newRowNumber = xlsx.utils.sheet_to_json(worksheet).length + 2; // Adjust for header
    xlsx.utils.sheet_add_aoa(worksheet, [newRow], { origin: `A${newRowNumber}` });
    xlsx.writeFile(workbook, registrationFilePath);

    res.send("Registration successful!");
});

// Route to handle general form data submission
app.post("/submit", (req, res) => {
    const formData = req.body;
    const workbook = getOrCreateWorkbook(dataFilePath, "Form Data", [
        'Name', 'Surname', 'Age', 'Phone Number', 'Proof Type',
        'Aadhar Number', 'Aadhar Photo', 'PAN Number', 'PAN Photo',
        'Joined By Name', 'Joined By Surname', 'Relation',
        'Joined By Email', 'Joined By Phone'
    ]);
    const worksheet = workbook.Sheets["Form Data"];

    const rowData = [
        formData.personalName,
        formData.personalSurname,
        formData.personalAge,
        formData.personalPhone,
        formData.proofType,
        formData.proofType === "aadhar" ? formData.aadharNumber : "",
        formData.proofType === "aadhar" ? formData.aadharPhoto : "",
        formData.proofType === "pan" ? formData.panNumber : "",
        formData.proofType === "pan" ? formData.panPhoto : "",
        formData.joinedName,
        formData.joinedSurname,
        formData.relation,
        formData.joinedEmail,
        formData.joinedPhone,
    ];

    const newRowNumber = xlsx.utils.sheet_to_json(worksheet).length + 2; // Adjust for header
    xlsx.utils.sheet_add_aoa(worksheet, [rowData], { origin: `A${newRowNumber}` });
    xlsx.writeFile(workbook, dataFilePath);

    res.status(200).json({ message: "Personal Information Form submitted successfully!" });
});

// Route to handle contact form data submission
app.post("/contact", (req, res) => {
    const { name, email, message } = req.body;
    const workbook = getOrCreateWorkbook(contactFilePath, "Contacts", ['Name', 'Email', 'Message', 'Date']);
    const worksheet = workbook.Sheets["Contacts"];

    const newEntry = [name, email, message, new Date().toLocaleString()];
    const newRowNumber = xlsx.utils.sheet_to_json(worksheet).length + 2; // Adjust for header
    xlsx.utils.sheet_add_aoa(worksheet, [newEntry], { origin: `A${newRowNumber}` });
    xlsx.writeFile(workbook, contactFilePath);

    res.json({ message: "Thank you for contacting" });
});

// Load and read the credentials.xlsx for login validation
const loadLoginData = () => {
    const workbook = xlsx.readFile('contact.xlsx');
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    return xlsx.utils.sheet_to_json(worksheet);
};

const loadLoginData1 = () => {
    const workbook = xlsx.readFile('registration.xlsx');
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    return xlsx.utils.sheet_to_json(worksheet);
};

// Endpoint to validate login
app.post('/validatelogin', (req, res) => {
    const { username, password } = req.body;
    const loginData = loadLoginData1();
    const isValid = loginData.some(record => record.Username === username && record.Password === password);

    if (isValid) {
        res.json({ success: true,data:loginData });
    } else {
        res.json({ success: false,data:loginData });
    }
});

// Endpoint to fetch the latest contact entry
app.get('/contact', (req, res) => {
    const contactData = getOrCreateWorkbook(contactFilePath, "Contacts", ['Name', 'Email', 'Message', 'Date']);
    const worksheet = contactData.Sheets["Contacts"];
    const contacts = xlsx.utils.sheet_to_json(worksheet);
   // const latestContact = contacts[contacts.length - 1]; // Get the last entry
    res.json(contacts || { message: "No contacts available" });
});


//Route to handle "Extra Services" form submissions
app.post('/extra-services', (req, res) => {
    const { name, email, food, bedsharing, accommodation, appliances, floor, environment } = req.body;
    
    // Set up workbook and worksheet
    const workbook = getOrCreateWorkbook(extraServicesFilePath, "Extra Services", ["Name", "Email", "Food", "Bed Sharing", "Accommodation", "Appliances", "Floor", "Environment"]);
    let worksheet = workbook.Sheets["Extra Services"];

    // Convert appliances array to a comma-separated string
    const appliancesString = appliances.join(", ");

    // Append new data to worksheet
    const data = [
        { Name: name, Email: email, Food: food, "Bed Sharing": bedsharing, Accommodation: accommodation, Appliances: appliancesString, Floor: floor, Environment: environment }
    ];
    const newWorksheet = xlsx.utils.sheet_add_json(worksheet, data, { origin: -1, skipHeader: true });
    workbook.Sheets["Extra Services"] = newWorksheet;

    // Write back to file
    xlsx.writeFile(workbook, extraServicesFilePath);
    res.json({ message: "Extra Services Form submitted successfully!" });
});


// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
