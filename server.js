const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const xlsx = require('xlsx');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(bodyParser.json());

const EXCEL_FILE = 'entries.xlsx';

// Load or create Excel file
const loadExcel = () => {
    if (!fs.existsSync(EXCEL_FILE)) {
        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.aoa_to_sheet([["PIS Number", "Name", "Computer IP", "Switch IP", "Port", "Building", "Room", "Domain", "Antivirus"]]);
        xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
        xlsx.writeFile(wb, EXCEL_FILE);
    }
    return xlsx.readFile(EXCEL_FILE);
};

// Search entries
app.get('/search', (req, res) => {
    const query = req.query.query.toLowerCase();
    const wb = loadExcel();
    const ws = wb.Sheets['Sheet1'];
    const data = xlsx.utils.sheet_to_json(ws, { header: 1 });
    const results = data.filter(row => row.some(cell => cell && cell.toString().toLowerCase().includes(query)));
    res.json(results.slice(1));
});

// Add new entry
app.post('/add_entry', (req, res) => {
    const entry = req.body;
    const wb = loadExcel();
    const ws = wb.Sheets['Sheet1'];
    const data = xlsx.utils.sheet_to_json(ws, { header: 1 });
    data.push([entry.pis_number, entry.name, entry.computer_ip, entry.switch_ip, entry.port, entry.building, entry.room, entry.domain, entry.antivirus]);
    ws['!ref'] = xlsx.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 8, r: data.length - 1 } });
    xlsx.utils.sheet_add_aoa(ws, data);
    xlsx.writeFile(wb, EXCEL_FILE);
    res.json({ status: 'Entry added successfully' });
});

// Update entry
app.put('/update_entry/:pis_number', (req, res) => {
    const pisNumber = req.params.pis_number;
    const updatedEntry = req.body;
    const wb = loadExcel();
    const ws = wb.Sheets['Sheet1'];
    const data = xlsx.utils.sheet_to_json(ws, { header: 1 });

    const index = data.findIndex(row => row[0] === pisNumber);
    if (index === -1) {
        return res.status(404).json({ error: 'Entry not found' });
    }

    data[index] = [updatedEntry.pis_number, updatedEntry.name, updatedEntry.computer_ip, updatedEntry.switch_ip, updatedEntry.port, updatedEntry.building, updatedEntry.room, updatedEntry.domain, updatedEntry.antivirus];
    ws['!ref'] = xlsx.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 8, r: data.length - 1 } });
    xlsx.utils.sheet_add_aoa(ws, data);
    xlsx.writeFile(wb, EXCEL_FILE);
    res.json({ status: 'Entry updated successfully' });
});

const PORT = 8000;
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});

