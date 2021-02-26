const fs = require('fs');
const readline = require('readline');
const {
  google
} = require('googleapis');
const {
  type
} = require('os');

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = 'token.json';
const idSpreedsheets = '1Z5fE9Bgu8BPkI7opjLjwjQZcVzt7oSTaL08QD8zkUuA';

fs.readFile('credentials.json', (err, content) => {
  if (err) return console.log('Error loading client secret file:', err);
  authorize(JSON.parse(content), getData);
  authorize(JSON.parse(content), updateSheet);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */

function authorize(credentials, callback) {
  const {
    client_secret,
    client_id,
    redirect_uris
  } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
    client_id, client_secret, redirect_uris[0]);

  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error while trying to retrieve access token', err);
      oAuth2Client.setCredentials(token);
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

/**
 * Prints the names and majors of students in a sample spreadsheet:
 * @see https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */


async function getData(auth) {
  const sheets = google.sheets({
    version: 'v4',
    auth
  });
  const allData = 'C4:H27';
  const result = await sheets.spreadsheets.values.get({
    spreadsheetId: idSpreedsheets,
    range: allData,
  });

  try {
    const rows = result.data.values;

    if (rows.length) {
      const rowsData = rows.map((row) => {
        return row;
      });

      console.log('getDataTable: ');
      console.log(rowsData);

      return checkStatus(auth, rowsData);

    } else {
      console.log('No data found.');
      return 0;
    }
  } catch (err) {
    if (err) return console.log('The API returned an error: ' + err);
  }
}

function checkStatus(auth, data) {
  console.log( 'Check status for student' );
  
  const dataUpdated = data.map(row => {
    console.log( ' Get student absences ' );
    let updated = reprovedByAbsence(row);

    console.log( ' Get student averange ' );
    const avg = calcAverange(row);

    console.log( ' Get stutent reproved ' );
    updated = reprovedByNote(avg, updated);

    console.log(' Get student approved ');
    updated = approved(avg, updated);

    console.log('Get student final exam ');
    updated = finalExam(avg, updated);
    
    console.log('Get note final');
    updated = calcFinalExam(avg, updated);

    return updated;
  })

  console.log('dataRules: ');
  console.log(dataUpdated);

  return updateSheet(auth, dataUpdated);
}

function reprovedByAbsence(row) {
  if (row[0] > 15) {
    row[4] = 'Reprovado por Falta';
    row[5] = '0';
  }

  return row;
}

function calcAverange(row) {
  const factorAvg = 3;
  const p1 = parseInt(row[1]);
  const p2 = parseInt(row[2]);
  const p3 = parseInt(row[3]);
  const averange = ((p1 + p2 + p3) / factorAvg);

  return Math.round(averange);
}

function reprovedByNote(avg, row) {
  const studentSituation = row[4];

  if (studentSituation) return row;

  if (avg < 50) {
    row[4] = 'Reprovado por Nota';
    row[5] = '0';
  }

  return row;
}

function approved(avg, row) {
  const studentSituation = row[4];

  if (studentSituation) return row;

  if (avg >= 70) {
    row[4] = 'Aprovado';
    row[5] = '0';
  }

  return row;
}

function finalExam(avg, row) {
  const studentSituation = row[4];

  if (studentSituation) return row;

  if ((avg >= 50) && (avg < 70)) {
    row[4] = 'Exame Final';
  }

  return row;
}

function calcFinalExam(avg, row) {
  let naf = row[5];

  if (naf === '0') return row;

  naf = parseInt(row[5]);

  const formNaf = (avg + naf) / 2;
  row[5] = formNaf;

  return row;
}

function updateSheet(auth, dataUpdated) {
  const sheets = google.sheets({
    version: 'v4',
    auth
  });

  const rangeUpdate = 'engenharia_de_software!C4:H27';

  sheets.spreadsheets.values.update({
    spreadsheetId: idSpreedsheets,
    range: rangeUpdate,
    resource: {
      range: rangeUpdate,
      majorDimension: 'ROWS',
      values: dataUpdated
    },
    valueInputOption: "USER_ENTERED",
  }, (err ) => {
    if (err) {
      console.log(err);
    }
  })
}