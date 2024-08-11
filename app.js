const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const app = express();

let highestBid = 0;

app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public'));

app.get('/', (req, res) => {
    res.render('index', { highestBid: highestBid });
});

app.post('/submit-bid', (req, res) => {
    const newBid = parseInt(req.body.bid);
    if (newBid > highestBid) {
        highestBid = newBid;
        res.redirect('/enter-details');
    } else {
        res.redirect('/');
    }
});

app.get('/enter-details', (req, res) => {
    res.render('details');
});

app.post('/save-details', (req, res) => {
    const { name, email } = req.body;
    const data = [
        {
            Name: name,
            Email: email,
            Bid: highestBid,
        },
    ];

    const file = '/tmp/bids.xlsx';  // Save to the temporary directory
    let workbook;
    let worksheet;

    if (fs.existsSync(file)) {
        workbook = xlsx.readFile(file);
        worksheet = workbook.Sheets['Bids'];
        if (!worksheet) {
            worksheet = xlsx.utils.json_to_sheet([]);
            xlsx.utils.book_append_sheet(workbook, worksheet, 'Bids');
        }
    } else {
        workbook = xlsx.utils.book_new();
        worksheet = xlsx.utils.json_to_sheet([]);
        xlsx.utils.book_append_sheet(workbook, worksheet, 'Bids');
    }

    xlsx.utils.sheet_add_json(worksheet, data, { skipHeader: true, origin: -1 });
    xlsx.writeFile(workbook, file);

    res.send('Details saved successfully! File is stored temporarily on the server.');
});

app.listen(3000, () => {
    console.log('Server is running on http://localhost:3000');
});

app.get('/download-bids', (req, res) => {
    const file = '/tmp/bids.xlsx'; // Path where your Excel file is stored

    if (fs.existsSync(file)) {
        res.download(file, 'bids.xlsx', (err) => {
            if (err) {
                res.status(500).send('Error downloading the file.');
            }
        });
    } else {
        res.status(404).send('File not found.');
    }
});
