const express = require('express');
const bodyParser = require('body-parser');
const excel = require('exceljs');
const path = require('path');

const cors = require('cors');

const app = express();

// Enable CORS for all routes
app.use(cors());

// Create a new Excel workbook and worksheet
const workbook = new excel.Workbook();
const worksheet = workbook.addWorksheet('Sheet1');

// Add column headers to the worksheet
worksheet.columns = [
  { header: 'Name', key: 'name', width: 20 },
  { header: 'Email', key: 'email', width: 30 },
  { header: 'Age', key: 'age', width: 10 },
];

// Use body-parser middleware to parse JSON requests
app.use(bodyParser.json());

// Define a route to handle form data submissions
app.post('/api/data', (req, res) => {
  const { name, email, age } = req.body;

  // Add a new row to the worksheet with the form data
  worksheet.addRow({ name, email, age });

  // Save the workbook to a file in the same directory as the server script
  const filePath = path.join(__dirname, 'data.xlsx');
  workbook.xlsx.writeFile(filePath)
    .then(() => {
      console.log('Data added to Excel file');
      console.log("File saved at:" + filePath);
      res.status(200).json({ message: "File saved at:" + filePath });
    })
    .catch((error) => {
      console.error(error);
      res.status(500).json({ error: 'Error adding data to Excel file' });
    });
});

// Define a route to download the Excel file
app.get('/api/data', (req, res) => {
  // Send the Excel file as a response to the GET request
  workbook.xlsx.writeBuffer()
    .then((buffer) => {
      res.setHeader('Content-Disposition', 'attachment; filename=data.xlsx');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(buffer);
    })
    .catch((error) => {
      console.error(error);
      res.status(500).json({ error: 'Error generating Excel file' });
    });
});

// Start the server
const port = 3000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});


//"File saved at:/opt/render/project/src/data.xlsx"
//Location of the saved excel file.