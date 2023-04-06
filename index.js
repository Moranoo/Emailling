const express = require('express');
const multer = require('multer');
const upload = multer({ dest: 'uploads/' });
const app = express();
const ExcelJS = require('exceljs');
const bodyParser = require('body-parser');
require('dotenv').config();


app.use(bodyParser.urlencoded({ extended: false }));

// Route pour afficher le formulaire de téléchargement de fichier Excel
app.get('/', (req, res) => {

  res.sendFile(__dirname + '/index.html');}
);

// Route pour afficher les 10 premiers mails
app.get('/emails', (req, res) => {
  const workbook = new ExcelJS.Workbook();
  workbook.xlsx.readFile('uploads/' + req.query.filename)
    .then(() => {
      const worksheet = workbook.getWorksheet(1);
      const column = worksheet.getColumn(3);
      const emails = column.values.slice(2, 12);
  
      res.send(`
        <h1>Les 10 premiers e-mails : </h1>
        <ul>
          ${emails.map(email => `<li>${email}</li>`).join('')}
        </ul>
        <form action="/send-email" method="POST">
          <label for="to">À :</label>
          <input type="email" id="to" name="to" required>
          <label for="subject">Sujet :</label>
          <input type="text" id="subject" name="subject" required>
          <label for="body">Corps du message :</label>
          <textarea id="body" name="body" required></textarea>
          <button type="submit">Envoyer</button>
        </form>
      `);
    })
    .catch(error => {
      console.log(`Erreur lors de la lecture du fichier Excel : ${error}`);
      res.send('Une erreur est survenue');
    });
});

// Route pour gérer le téléchargement d'un fichier Excel
app.post('/upload', upload.single('excelFile'), (req, res) => {
  // Initialiser le message d'erreur à une chaîne vide
  let errorMsg = '';

  // Vérifier si un fichier a été chargé
  if (!req.file) {
    errorMsg = 'Veuillez sélectionner un fichier Excel';
  } else {
    // Lecture du fichier Excel
    const workbook = new ExcelJS.Workbook();
    workbook.xlsx.readFile('uploads/' + req.file.filename)
      .then(() => {
        const worksheet = workbook.getWorksheet(1);
        const column = worksheet.getColumn(3);
        const emails = column.values.slice(2);

        console.log(emails);
      })
      .catch(error => {
        console.log(`Erreur lors de la lecture du fichier Excel : ${error}`);
      });

    // Redirection vers la page des e-mails
    res.redirect(`/emails?filename=${req.file.filename}`);
  }

  // Afficher la page HTML avec le message d'erreur éventuel
  res.send(`
    <h1>Téléchargement de fichier Excel</h1>
    <form method="post" action="/upload" enctype="multipart/form-data">
      <input type="file" name="excelFile" accept=".xlsx, .xls, .csv"><br>
      <span style="color: red;">${errorMsg}</span>
      <br>
      <button type="submit">Envoyer</button>
    </form>
    <form action="/" method="GET">
      <button type="submit">Retour</button>
    </form>
    `);
});

app.post('/send-email', (req, res) => {
  const { to, subject, body } = req.body;

  const nodemailer = require('nodemailer');
  const transporter = nodemailer.createTransport({
    service: 'outlook',
    auth: {
        user: process.env.EMAIL,
        pass: process.env.PASSWORD
    }
  });


  const mailOptions = {
    from: process.env.EMAIL, // Adresse e-mail de l'expéditeur
    to, // Adresse e-mail du destinataire
    subject, // Sujet du mail
    html: body // Corps du mail (peut contenir du HTML)
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.log(`Erreur lors de l'envoi de l'e-mail : ${error}`);
      res.send('Une erreur est survenue');
    } else {
      console.log(`E-mail envoyé avec succès : ${info.response}`);
      res.send('E-mail envoyé avec succès');
    }
  });
});


app.listen(3000, () => {
  console.log('Serveur démarré sur le port 3000');
});
