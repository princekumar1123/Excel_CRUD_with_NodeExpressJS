const express = require('express');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const app = express();
const PORT = 9000;
const excelFilePath = path.join(__dirname, 'data.xlsx');

app.use(express.json());

function readExcelData() {
    if (!fs.existsSync(excelFilePath)) {
        const wb = xlsx.utils.book_new();
        const ws = xlsx.utils.json_to_sheet([]);
        xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
        xlsx.writeFile(wb, excelFilePath);
    }
    const workbook = xlsx.readFile(excelFilePath);
    const worksheet = workbook.Sheets['Sheet1'];
    return xlsx.utils.sheet_to_json(worksheet);
}

function writeExcelData(data) {
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet(data);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    xlsx.writeFile(workbook, excelFilePath);
}


app.get('/records', (req, res) => {
    const data = readExcelData();
    console.log("data", data);

    res.json(data);
});

app.post('/records', (req, res) => {
    const idGenerator = () => {
        return Math.floor(Math.random() * 1_000_000_000)
    }
    const data = readExcelData();
    const newRecord = req.body;
    const currentID = idGenerator()
    const existIDs = data.map((e, i) => e.id)    
    if (existIDs.includes(currentID)) {
        currentID = idGenerator()
    }
    data.push({ ...newRecord, id: currentID });
    writeExcelData(data);
    res.status(201).json({ message: 'Record added successfully!', record: newRecord });
});

app.put('/records/:id', (req, res) => {
    const data = readExcelData();
    const { id } = req.params;
    const updatedRecord = req.body;

    const recordIndex = data.findIndex(record => record.id === JSON.parse(id));
    console.log("recordIndex", data, recordIndex);

    if (recordIndex === -1) {
        return res.status(404).json({ message: 'Record not found!' });
    }

    data[recordIndex] = { ...data[recordIndex], ...updatedRecord };
    writeExcelData(data);
    res.json({ message: 'Record updated successfully!', record: data[recordIndex], status: true });
});

app.delete('/records/:id', (req, res) => {
    const data = readExcelData();
    const { id } = req.params;

    const newData = data.filter(record => record.id != id);
    if (data.length === newData.length) {
        return res.status(404).json({ message: 'Record not found!' });
    }

    writeExcelData(newData);
    res.json({ message: 'Record deleted successfully!' });
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
