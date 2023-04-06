const nodemailer = require('nodemailer');

async function sendEmail(to, subject, text) {
  const transporter = nodemailer.createTransport({
    service: 'outlook',
    auth: {
      user: 'myspeciality@outlook.com',
      pass: ''
    }
  });

  const mailOptions = {
    from: 'myspeciality@outlook.com',
    to,
    subject: 'CPLF',
    html: `
    <html>
    <head>
      <style>
        body {
          background-color: #f2f2f2;
        }
      </style>
    </head>
    <body>
      <h1>Voici le compte rendu</h1>
    
      <a href="https://www.google.com/">Interview de billy</a>
      <img src="logo/Logo MySpec.png"  alt="Texte alternatif">
      </body>
      </html>
    `
  };

  const info = await transporter.sendMail(mailOptions);
  console.log('E-mail envoyé avec succès:', info.response);
}
module.exports = {
  sendEmail: sendEmail
};