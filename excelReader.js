const xlsx = require('xlsx');

module.exports = function readEmailsFromExcel(filePath) {
  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];

  const emails = [];
  for (let i = 2; i <= worksheet['!ref'].split(':')[1][1]; i++) {
    const email = worksheet[`A${i}`].v.trim();
    if (email && email.match(/\S+@\S+\.\S+/)) {
      emails.push(email);
    }
  }

  return emails;
}
