// app.js
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const fs = require('fs');
const xlsx = require('xlsx');

const app = express();
app.use(bodyParser.json());
app.use(cors());

// Load the Excel file
const filePath = './data.xlsx'; // Excel file path
let workbook;

// Helper function to read Excel file
function loadWorkbook() {
    if (fs.existsSync(filePath)) {
        workbook = xlsx.readFile(filePath);
    } else {
        workbook = xlsx.utils.book_new();
        workbook.SheetNames.push("Sheet1");
        workbook.Sheets["Sheet1"] = xlsx.utils.json_to_sheet([]);
    }
}

// Load workbook initially
loadWorkbook();

// Search entries by PIS number
app.get('/search', (req, res) => {
    const { pis } = req.query;
    const sheet = workbook.Sheets["Sheet1"];
    const rows = xlsx.utils.sheet_to_json(sheet);

    // Filter entries by PIS number
    const results = rows.filter(row => row.PIS === pis);
    res.json(results);
});

// Add a new entry
app.post('/add', (req, res) => {
    const newEntry = req.body;
    const sheet = workbook.Sheets["Sheet1"];
    const rows = xlsx.utils.sheet_to_json(sheet);

    // Add the new entry to rows
    rows.push(newEntry);

    // Write back to the Excel file
    const updatedSheet = xlsx.utils.json_to_sheet(rows);
    workbook.Sheets["Sheet1"] = updatedSheet;
    xlsx.writeFile(workbook, filePath);
    res.json({ success: true, message: "Entry added successfully!" });
});

// Update an existing entry
app.put('/update', (req, res) => {
    const updatedEntry = req.body;
    const sheet = workbook.Sheets["Sheet1"];
    const rows = xlsx.utils.sheet_to_json(sheet);

    // Find the entry to update by PIS and IP or Port
    const index = rows.findIndex(row => row.PIS === updatedEntry.PIS && row['Computer IP'] === updatedEntry['Computer IP']);
    if (index > -1) {
        rows[index] = updatedEntry;

        // Write updated data back to the Excel file
        const updatedSheet = xlsx.utils.json_to_sheet(rows);
        workbook.Sheets["Sheet1"] = updatedSheet;
        xlsx.writeFile(workbook, filePath);
        res.json({ success: true, message: "Entry updated successfully!" });
    } else {
        res.status(404).json({ success: false, message: "Entry not found!" });
    }
});

// Run the server
const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});

