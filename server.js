const express = require('express');
const bodyParser = require('body-parser');
const fs = require('fs');
const XLSX = require('xlsx');
const app = express();
const port = 3000;

// Middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Endpoint to handle form submission
app.post('/submit-registration', (req, res) => {
  const data = req.body;

  // Convert data to array format
  const row = [
    data.fullName,
    data.fatherName,
    data.motherName,
    data.birthDate,
    data.address,
    data.phone,
    data.email,
    data.education,
    data.course,
    data.chronicDisease,
    data.diseaseDetails
  ];

  // Read existing Excel file or create a new one
  let workbook;
  const filePath = 'registrations.xlsx';
  if (fs.existsSync(filePath)) {
    workbook = XLSX.readFile(filePath);
  } else {
    workbook = XLSX.utils.book_new();
    workbook.SheetNames.push('Registrations');
    const worksheet = XLSX.utils.aoa_to_sheet([[
      'Full Name', 'Father Name', 'Mother Name', 'Birth Date', 'Address', 'Phone', 'Email',
      'Education', 'Course', 'Chronic Disease', 'Disease Details'
    ]]);
    workbook.Sheets['Registrations'] = worksheet;
  }

  // Append new data
  const worksheet = workbook.Sheets['Registrations'];
  const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  sheetData.push(row);
  const newWorksheet = XLSX.utils.aoa_to_sheet(sheetData);
  workbook.Sheets['Registrations'] = newWorksheet;

  // Write updated workbook to file
  XLSX.writeFile(workbook, filePath);

  res.send('Registration successful');
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
