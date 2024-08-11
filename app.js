const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const session = require('express-session');
const app = express();

let highestBid = 0;
let bidClosed = false;

// Admin credentials
const adminUsername = "admin";
let adminPassword = "password"; // Change this to a strong password

app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static('public')); // Serve static files from the 'public' folder

app.use(session({
    secret: 'your-secret-key',
    resave: false,
    saveUninitialized: true,
}));

// Home Page
app.get('/', (req, res) => {
    if (bidClosed) {
        res.send("The bid is closed.");
    } else {
        res.render('index', { highestBid: highestBid });
    }
});

// Submit Bid
app.post('/submit-bid', (req, res) => {
    if (bidClosed) {
        res.send("The bid is closed.");
        return;
    }

    const newBid = parseInt(req.body.bid);
    const college = req.body.college; // Get the college name from the form

    if (newBid > highestBid) {
        highestBid = newBid;

        // Save the bid and college to the Excel file
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

        // Add headers if not present
        if (!worksheet['A1']) {
            worksheet['A1'] = { v: 'Bid', t: 's' };
            worksheet['B1'] = { v: 'College', t: 's' };
        }

        // Add new data including the college name
        const data = [{
            Bid: highestBid,
            College: college
        }];

        xlsx.utils.sheet_add_json(worksheet, data, { skipHeader: true, origin: -1 });
        xlsx.writeFile(workbook, file);

        res.redirect('/enter-details');
    } else {
        res.redirect('/');
    }
});

// Enter Details Page
app.get('/enter-details', (req, res) => {
    res.render('details'); // Ensure 'details.ejs' exists in the views folder
});

// Save Details and Update Bid
app.post('/save-details', (req, res) => {
    const { name, email, whatsapp } = req.body; // Extract whatsapp from req.body
    const emailPattern = /^[a-zA-Z]+\.[a-zA-Z]+\.MBA23@said\.oxford\.edu$/;
    
    if (!emailPattern.test(email)) {
        return res.send('Enter your SBS email');
    }

    // Check if both name and email are provided
    if (name && email) {
        const data = [
            {
                Name: name,
                Email: email,
                WhatsApp: whatsapp, // Include WhatsApp number in the data
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

        // Add headers if not present
        if (!worksheet['A1']) {
            worksheet['A1'] = { v: 'Name', t: 's' };
            worksheet['B1'] = { v: 'Email', t: 's' };
            worksheet['C1'] = { v: 'WhatsApp', t: 's' };
            worksheet['D1'] = { v: 'Bid', t: 's' };
        }

        // Add new data
        xlsx.utils.sheet_add_json(worksheet, data, { skipHeader: true, origin: -1 });
        xlsx.writeFile(workbook, file);

        res.send(`
            <html>
            <head>
                <title>Thank You</title>
                <link rel="stylesheet" href="/styles.css">
            </head>
            <body>
                <h1>Thank you for showing your interest.</h1>
                <p>Your details have been saved successfully. We will contact you soon.</p>
                <a href="/">Return to Home Page</a>
            </body>
            </html>
        `);
    } else {
        res.send('Please provide both name and email.');
    }
});

// Admin Login Page
app.get('/admin-login', (req, res) => {
    res.render('admin/admin-login'); // Correct path for admin-login.ejs
});

// Admin Login Handler
app.post('/admin-login', (req, res) => {
    const { username, password } = req.body;
    if (username === adminUsername && password === adminPassword) {
        req.session.isAdmin = true;
        res.redirect('/admin');
    } else {
        res.send('Invalid credentials. Please try again.');
    }
});

// Middleware to protect admin routes
function ensureAdmin(req, res, next) {
    if (req.session.isAdmin) {
        next();
    } else {
        res.redirect('/admin-login');
    }
}

// Admin Dashboard
app.get('/admin', ensureAdmin, (req, res) => {
    res.render('admin/admin-dashboard', { highestBid }); // Correct path for admin-dashboard.ejs
});

// Close Bid
app.post('/admin/close-bid', ensureAdmin, (req, res) => {
    bidClosed = true;
    res.redirect('/admin');
});

// Reset Bid
app.post('/admin/reset-bid', ensureAdmin, (req, res) => {
    highestBid = 0;
    bidClosed = false;

    // Rename the existing file to avoid overwriting
    const oldFile = './bids.xlsx';
    if (fs.existsSync(oldFile)) {
        const newFileName = `./bids_${Date.now()}.xlsx`;
        fs.renameSync(oldFile, newFileName);
    }

    res.redirect('/admin');
});

// Download Excel Sheet
app.get('/admin/download-excel', ensureAdmin, (req, res) => {
    const file = path.resolve(__dirname, 'bids.xlsx');
    res.download(file);
});

// Change Password
app.post('/admin/change-password', ensureAdmin, (req, res) => {
    const { oldPassword, newPassword } = req.body;
    if (oldPassword === adminPassword) {
        adminPassword = newPassword; // Update the password
        res.redirect('/admin'); // Redirect to admin dashboard after update
    } else {
        res.send('Incorrect old password. Please try again.');
    }
});

// Admin Logout
app.post('/admin-logout', (req, res) => {
    req.session.destroy(err => {
        if (err) {
            return res.send('Error logging out.');
        }
        res.redirect('/admin-login');
    });
});

app.listen(3000, () => {
    console.log('Server is running on http://localhost:3000');
});
