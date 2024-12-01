// Prepare a google sheet with 3 columns: emailaddress, name, secretsanta(empty)

function assignRandomGiveto() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();

  const headerRow = 1;
  const emailColumnIndex = 0;
  const nameColumnIndex = 1;
  const secretSantaColumnIndex = 2;

  for (let i = headerRow; i < data.length; i++) {
    data[i][secretSantaColumnIndex] = "";
  }

  const names = data.slice(headerRow).map(row => row[nameColumnIndex]);
  const shuffledNames = shuffleArray([...names]);

  for (let i = 0; i < names.length; i++) {
    let assignedName = shuffledNames[i];
    if (assignedName === names[i]) {
      const swapIndex = (i + 1) % names.length;
      [shuffledNames[i], shuffledNames[swapIndex]] = [shuffledNames[swapIndex], shuffledNames[i]];
      assignedName = shuffledNames[i];
    }
    data[i + headerRow][secretSantaColumnIndex] = assignedName;

    // Send email notifications
    const emailAddress = data[i + headerRow][emailColumnIndex];
    const recipientName = data[i + headerRow][nameColumnIndex];
    const subject = "secret santa notification";
    const message = `merry christmas ${recipientName}, your secret santa is ${assignedName}`;
    MailApp.sendEmail(emailAddress, subject, message);
  }

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  const backgroundRange = sheet.getRange(headerRow + 1, secretSantaColumnIndex + 1, data.length - headerRow, 1);
  backgroundRange.setBackground("black");
}

function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
  return array;
}
