const express = require('express');
const serverless = require('serverless-http');
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage() });
const bodyParser = require('body-parser');
const { google } = require('googleapis');
const path = require('path');
const { uploadFile } = require('./driveUploader');

const app = express();
const SPREADSHEET_ID = '1S8oHwZ839_cfFq1o1dK82HW9fCScntawCqX1zXPy15k';
const SHEET_NAME = 'Dataset';

// View Engine Setup
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, '..', 'views'));
app.use(bodyParser.urlencoded({ extended: true, limit: '1mb' }));
app.use(bodyParser.json({ limit: '1mb' }));

// Google Sheets Auth
const auth = new google.auth.GoogleAuth({
  credentials: {
    type: 'service_account',
    project_id: process.env.GOOGLE_PROJECT_ID,
    private_key_id: process.env.GOOGLE_PRIVATE_KEY_ID,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    client_id: process.env.GOOGLE_CLIENT_ID,
    auth_uri: process.env.GOOGLE_AUTH_URI,
    token_uri: process.env.GOOGLE_TOKEN_URI,
    auth_provider_x509_cert_url: process.env.GOOGLE_AUTH_PROVIDER_X509_CERT_URL,
    client_x509_cert_url: process.env.GOOGLE_CLIENT_X509_CERT_URL,
  },
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
  ]
});

const sheets = google.sheets({ version: 'v4', auth });

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, '..', 'views'));

const session = require('express-session');

app.use(session({
  secret: 'Xnxx1230987azkaArga', // ganti dengan yang lebih aman
  resave: false,
  saveUninitialized: true
}));


function requireLogin(req, res, next) {
  if (req.session && req.session.loggedIn) {
    next();
  } else {
    res.redirect('/login');
  }
}


// Middleware //
app.use(bodyParser.urlencoded({ extended: true, limit: '1mb' }));
app.use(bodyParser.json({ limit: '1mb' }));

async function fetchDataFromGoogleSheet() {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1:BK`, // Ambil data secara langsung dari Google Sheets
    });
    return response.data.values.slice(2) || []; // Exclude header
  } catch (error) {
    console.error('Error fetching data from Google Sheets:', error);
    throw error;
  }
}

// Data Drodown//
async function fetchDropdownList() {
  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Dropdown!A2:H15',
    });
    return response.data.values || [];
  } catch (error) {
    console.error('Error fetching dropdown list:', error);
    throw error;
  }
}


app.get('/login', (req, res) => {
  res.render('login', { error: null });
});

app.post('/login', async (req, res) => {
  const { email, password } = req.body;

  try {
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Dropdown!L2:N15',
    });

    const users = response.data.values || [];

    const userFound = users.find(row =>
      row[0]?.toLowerCase() === email.toLowerCase() &&
      row[1] === password
    );

if (userFound) {
  req.session.loggedIn = true;
  req.session.email = email;
  req.session.userName = userFound[2]; // Nama user dari kolom M
  req.session.loginTime = Date.now();  // ⬅️ Tambahkan ini
  res.redirect('/data');
    } else {
      res.render('login', { error: 'Email atau Password salah' });
    }
  } catch (error) {
    console.error('Login error:', error);
    res.render('login', { error: 'Terjadi kesalahan saat login' });
  }
});



// Route to display data in a table
app.get('/data', requireLogin, async (req, res) => {
  try {
    const selectedProduct = req.query.main_product || '';
    const data = await fetchDataFromGoogleSheet();

  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  const filteredData = data.filter(row => {
  const productMatch = !selectedProduct || (row[0] || '').toLowerCase() === selectedProduct.toLowerCase();
  const billDate = row[10]; // kolom ke-10 (index ke-9)
  
  if (!billDate || typeof billDate !== 'string') return false;

  const parsedDate = new Date(billDate);
  return (
    productMatch &&
    !isNaN(parsedDate) &&
    parsedDate.getMonth() === currentMonth &&
    parsedDate.getFullYear() === currentYear
  );
});

res.render('dataTable', {
  data: filteredData,
  selectedProduct,
  userName: req.session.userName,
  loginTime: req.session.loginTime // ✅ tambahkan ini
});

  } catch (error) {
    console.error('Error fetching data:', error);
    res.status(500).send('Error loading data');
  }
});


// Route to display form for adding new data
app.get('/add', requireLogin, async (req, res) => {
  try {
    if (req.query.noCache === 'true') {
      dataCache = null;
      cacheTimestamp = null;
    }

    const listData = await fetchDropdownList();
    const userOptions = listData.map(row => row[0]).filter(item => item && item.trim() !== '');
    const messageOptions = listData.map(row => row[1]).filter(item => item && item.trim() !== '');
    const callOptions = listData.map(row => row[2]).filter(item => item && item.trim() !== '');
    const suratOptions = listData.map(row => row[3]).filter(item => item && item.trim() !== '');
    const visitOptions = listData.map(row => row[4]).filter(item => item && item.trim() !== '');

    res.render('form', {
      editMode: false,
      userOptions,
      messageOptions,
      callOptions,
      suratOptions,
      visitOptions
    });
  } catch (error) {
    console.error('Error fetching dropdown data:', error);
    res.status(500).send('Error loading form');
  }
});

// Route to handle form submission
app.post('/submit', requireLogin, async (req, res) => {
  const { main_product, agrrement_number, debitur_name, debt_phone, alamat_debt, econ_name, econ_phone, tenor, current_period,  installment, bill_date, due_date, overdue, 
          repaid_amount, os_principal, status_debtur, 
          user_name_sp, 
          sp1_date, sp1_no, sp1_send, sp1_status, sp2_date, sp2_no, sp2_send, sp2_status, sp3_date, sp3_no, sp3_send, sp3_status, 
          user_name_wa,
          wa_date, wa_status, wa_kronologi, wa_tmp,
          user_name_call, 
          call1_debt_date, call1_debt_status, call1_debt_tmp, 
          call2_debt_date, call2_debt_status, call2_debt_tmp, call_debt_kronologi,
          call1_econ_date, call1_econ_status, call1_econ_tmp, 
          call2_econ_date, call2_econ_status, call2_econ_tmp,call_econ_kronologi,
          new_contact, call_other_date, call_other_status, call_other_tmp, call_other_kronologi,
          user_name_visit, 
          visit_date, visit_Poto, visit_kronologi, visit_location, janji_bayar, action_plan,
          sk_date, result_external} = req.body;
          try {
    await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: SHEET_NAME,
    valueInputOption: 'USER_ENTERED',
    resource: {
    values: [[main_product, agrrement_number, debitur_name, debt_phone, alamat_debt, econ_name, econ_phone, tenor, current_period,  installment, bill_date, '', '', 
          '', '', '', 
          '', 
          '', '', '', sp1_status, '', '', '', sp2_status, '', '', '', sp3_status, 
          user_name_wa,
          wa_date, wa_status, wa_kronologi, wa_tmp,
          user_name_call, 
          call1_debt_date, call1_debt_status, call1_debt_tmp, 
          call2_debt_date, call2_debt_status, call2_debt_tmp, call_debt_kronologi,
          call1_econ_date, call1_econ_status, call1_econ_tmp, 
          call2_econ_date, call2_econ_status, call2_econ_tmp,call_econ_kronologi,
          new_contact, call_other_date, call_other_status, call_other_tmp, call_other_kronologi,
          user_name_visit, 
          visit_date, visit_Poto, visit_kronologi, visit_location, janji_bayar, action_plan,
          sk_date, result_external]],
      },
    });
    // Clear cache
    dataCache = null;
    res.redirect('/data');
  } catch (error) {
    console.error('Error adding data:', error);
    res.status(500).send('Error adding data');
  }
});

// Route to display data in edit form
app.get('/edit-id/:main_product/:agreement_number', requireLogin, async (req, res) => {
  const { main_product, agreement_number } = req.params;

  try {
    const data = await fetchDataFromGoogleSheet();
    const row = data.find(r => r[0] === main_product && r[1] === agreement_number);

    if (!row) {
      return res.status(404).send('Data tidak ditemukan');
    }

    const listData = await fetchDropdownList();
    const userOptions = listData.map(r => r[0]).filter(item => item?.trim());
    const messageOptions = listData.map(r => r[1]).filter(item => item?.trim());
    const callOptions = listData.map(r => r[2]).filter(item => item?.trim());
    const suratOptions = listData.map(r => r[3]).filter(item => item?.trim());
    const visitOptions = listData.map(r => r[4]).filter(item => item?.trim());

    res.render('form', {
      row,
      editMode: true,
      userOptions,
      messageOptions,
      callOptions,
      suratOptions,
      visitOptions
    });
  } catch (error) {
    console.error('Error loading edit form:', error);
    res.status(500).send('Error saat edit');
  }
});




// Route to handle update
app.post('/update/:main_product/:agreement_number', upload.array('visit_Poto'), requireLogin, async (req, res) => {
  const { main_product, agreement_number } = req.params;
  const body = req.body;

  try {
    const data = await fetchDataFromGoogleSheet();
    const index = data.findIndex(row => row[0] === main_product && row[1] === agreement_number);
    if (index === -1) {
      return res.status(404).send('Data tidak ditemukan');
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `Dataset!A${index + 3}:BK${index + 3}`,
    });

    const oldRow = response.data.values[0] || [];
    const oldVisitPoto = oldRow[56] || '';
    const oldLinkArray = oldVisitPoto ? oldVisitPoto.split('\n') : [];

    let newLinks = [];
    if (req.files && req.files.length > 0) {
      for (const file of req.files) {
        const link = await uploadFile(file);
        newLinks.push(link);
      }
    }

    const newLinkArray = newLinks.map((link, i) => {
      const total = oldLinkArray.length + i + 1;
      return `[Foto ${total}](${link})`;
    });
    const visit_Poto = [...oldLinkArray, ...newLinkArray].join('\n');

    const values = [
      body.main_product, body.agrrement_number, body.debitur_name, body.debt_phone, body.alamat_debt, body.econ_name, body.econ_phone, body.tenor,
      body.current_period, body.installment, body.bill_date, '', '', '', '', '',
      '', '', '', '', body.sp1_status, '', '', '', body.sp2_status, 
      '', '', '', body.sp3_status, 
      body.user_name_wa, body.wa_date, body.wa_status, body.wa_kronologi, body.wa_tmp,
      body.user_name_call, body.call1_debt_date, body.call1_debt_status, body.call1_debt_tmp,
      body.call2_debt_date, body.call2_debt_status, body.call2_debt_tmp, body.call_debt_kronologi,
      body.call1_econ_date, body.call1_econ_status, body.call1_econ_tmp,
      body.call2_econ_date, body.call2_econ_status, body.call2_econ_tmp, body.call_econ_kronologi,
      body.new_contact, body.call_other_date, body.call_other_status, body.call_other_tmp, body.call_other_kronologi,
      body.user_name_visit, body.visit_date, visit_Poto, body.visit_kronologi, body.visit_location, body.janji_bayar, body.action_plan,
      body.sk_date, body.result_external
    ];

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Dataset!A${index + 3}:BK${index + 3}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [values] },
    });

    res.redirect('/data');
  } catch (error) {
    console.error('Error updating data:', error);
    res.status(500).send('Error updating data');
  }
});




app.get('/clear-cache', (req, res) => {
  dataCache = null;
  cacheTimestamp = null;
  res.send('Cache cleared!');
});

app.get('/logout', (req, res) => {
  req.session.destroy((err) => {
    if (err) {
      console.error('Logout error:', err);
    }
    res.redirect('/login');
  });
});


app.get('/', (req, res) => {
  res.send('Aplikasi Productivity sudah running di Vercel!');
});

module.exports = app;
module.exports.handler = serverless(app);
