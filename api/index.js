const express = require('express');
const serverless = require('serverless-http');
const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage() });
const bodyParser = require('body-parser');
const { google } = require('googleapis');
const path = require('path');
const { uploadFile } = require('./driveUploader');

const app = express();
const SPREADSHEET_ID = '1T3ti9E3PrUY6LvxFKE1Q_4VjzXABzi3dpgDRCqwpH1s';
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
      range: `${SHEET_NAME}!A1:BL`, // Ambil data secara langsung dari Google Sheets
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
      visitOptions,
      userName: req.session.userName //baru pasang
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
          visit_date, visit_ke, visit_Poto, visit_kronologi, visit_location, janji_bayar, action_plan,
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
          visit_date, visit_ke, visit_Poto, visit_kronologi, visit_location, janji_bayar, action_plan,
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
      visitOptions,
      userName: req.session.userName 
    });
  } catch (error) {
    console.error('Error loading edit form:', error);
    res.status(500).send('Error saat edit');
  }
});




// Route to handle update
// Route to handle update
// Route to handle update
app.post('/update/:main_product/:agreement_number', upload.array('visit_Poto'), requireLogin, async (req, res) => {
  const { main_product, agreement_number } = req.params;
  const body = req.body;

  function safeValue(newVal, oldVal) {
    return newVal && newVal.trim() !== "" ? newVal : (oldVal || "");
  }

  try {
    const data = await fetchDataFromGoogleSheet();

    let rowsToUpdate = [];

    if (main_product.toLowerCase() === 'heavy equipment') {
      // Cari nama debitur dari row yang sedang diupdate
      const targetRow = data.find(r => r[0] === main_product && r[1] === agreement_number);
      if (!targetRow) return res.status(404).send('Data tidak ditemukan');

      const targetDebitur = targetRow[2];
      rowsToUpdate = data
        .map((row, idx) => ({ row, idx }))
        .filter(item => item.row[2] === targetDebitur);

    } else {
      const index = data.findIndex(row => row[0] === main_product && row[1] === agreement_number);
      if (index !== -1) {
        rowsToUpdate.push({ row: data[index], idx: index });
      }
    }

    if (rowsToUpdate.length === 0) {
      return res.status(404).send('Data tidak ditemukan');
    }

    for (const { idx } of rowsToUpdate) {
      // Ambil data lama
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: `Dataset!A${idx + 3}:BL${idx + 3}`,
      });
      const oldRow = response.data.values[0] || [];
      const oldVisitPoto = oldRow[57] || '';
      const oldLinkArray = oldVisitPoto ? oldVisitPoto.split('\n') : [];

      // Upload foto baru (jika ada)
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

      // Susun data dengan merge data lama
      const values = [
        safeValue(body.main_product, oldRow[0]),
        safeValue(body.agrrement_number, oldRow[1]),
        safeValue(body.debitur_name, oldRow[2]),
        safeValue(body.debt_phone, oldRow[3]),
        safeValue(body.alamat_debt, oldRow[4]),
        safeValue(body.econ_name, oldRow[5]),
        safeValue(body.econ_phone, oldRow[6]),
        safeValue(body.tenor, oldRow[7]),
        safeValue(body.current_period, oldRow[8]),
        safeValue(body.installment, oldRow[9]),
        safeValue(body.bill_date, oldRow[10]),

        safeValue("", oldRow[11]),
        safeValue("", oldRow[12]),
        safeValue("", oldRow[13]),
        safeValue("", oldRow[14]),
        safeValue("", oldRow[15]),
        safeValue("", oldRow[16]),
        safeValue("", oldRow[17]),
        safeValue("", oldRow[18]),
        safeValue("", oldRow[19]),

        safeValue(body.sp1_status, oldRow[20]),
        safeValue("", oldRow[21]),
        safeValue("", oldRow[22]),
        safeValue("", oldRow[23]),

        safeValue(body.sp2_status, oldRow[24]),
        safeValue("", oldRow[25]),
        safeValue("", oldRow[26]),
        safeValue("", oldRow[27]),        

        safeValue(body.sp3_status, oldRow[28]),
        safeValue(body.user_name_wa, oldRow[29]),
        safeValue(body.wa_date, oldRow[30]),
        safeValue(body.wa_status, oldRow[31]),
        safeValue(body.wa_kronologi, oldRow[32]),
        safeValue(body.wa_tmp, oldRow[33]),
        safeValue(body.user_name_call, oldRow[34]),
        safeValue(body.call1_debt_date, oldRow[35]),
        safeValue(body.call1_debt_status, oldRow[36]),
        safeValue(body.call1_debt_tmp, oldRow[37]),
        safeValue(body.call2_debt_date, oldRow[38]),
        safeValue(body.call2_debt_status, oldRow[39]),
        safeValue(body.call2_debt_tmp, oldRow[40]),
        safeValue(body.call_debt_kronologi, oldRow[41]),
        safeValue(body.call1_econ_date, oldRow[42]),
        safeValue(body.call1_econ_status, oldRow[43]),
        safeValue(body.call1_econ_tmp, oldRow[44]),
        safeValue(body.call2_econ_date, oldRow[45]),
        safeValue(body.call2_econ_status, oldRow[46]),
        safeValue(body.call2_econ_tmp, oldRow[47]),
        safeValue(body.call_econ_kronologi, oldRow[48]),
        safeValue(body.new_contact, oldRow[49]),
        safeValue(body.call_other_date, oldRow[50]),
        safeValue(body.call_other_status, oldRow[51]),
        safeValue(body.call_other_tmp, oldRow[52]),
        safeValue(body.call_other_kronologi, oldRow[53]),
        safeValue(body.user_name_visit, oldRow[54]),
        safeValue(body.visit_date, oldRow[55]),
        safeValue(body.visit_ke, oldRow[56]),
        visit_Poto, // kolom foto gabungan lama + baru
        safeValue(body.visit_kronologi, oldRow[58]),
        safeValue(body.visit_location, oldRow[59]),
        safeValue(body.janji_bayar, oldRow[60]),
        safeValue(body.action_plan, oldRow[61]),
        safeValue(body.sk_date, oldRow[62]),
        safeValue(body.result_external, oldRow[63])
      ];

      // Update ke Google Sheets
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `Dataset!A${idx + 3}:BL${idx + 3}`,
        valueInputOption: 'USER_ENTERED',
        resource: { values: [values] },
      });
    }

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
