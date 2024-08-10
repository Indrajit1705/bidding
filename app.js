const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const app = express();

let highestBid = 0;

app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public')); // Serve static files from the 'public' folder

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

    const file = './bids.xlsx';
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

    res.send('Details saved successfully!');
});

app.listen(3000, () => {
    console.log('Server is running on http://localhost:3000');
});
