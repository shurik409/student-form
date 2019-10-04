const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');
const mkdirp = require('mkdirp');
const path = require('path');
const rimraf = require("rimraf");
const fs = require('fs');

const url = require('url');
const { app, BrowserWindow } = require('electron');


const { google } = require('googleapis');
const fetch = require('node-fetch');

const creds = require('./secret2.json');
const faculties = new Map();

async function accessSpreadsheet() {
    console.log(1);
    const doc = new GoogleSpreadsheet('1kCK4YoxAvDCkiq6f547Osyy0LWFmmdRB947XfuytiQM');
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)()
    const sheet = info.worksheets[0];

    const rows = await promisify(sheet.getRows)({
        offset: 1
    });
    rows.forEach(row => {
        let student = parseStudent(row);
        student.faculty.forEach(fac => {
            if (faculties.get(fac.trim())){
                faculties.get(fac.trim()).push(student);
            } else {
                faculties.set(fac.trim(), [ student ])
            }
        })
    })
    await rimraf.sync(path.join(__dirname, 'Факультеты'));
    mkdirp(path.join(__dirname, `Факультеты`))
    mkdirp(path.join(__dirname, `Test`))
    faculties.forEach((students, fac) => {
        mkdirp(path.join(__dirname, `Факультеты/${fac}`))
        students.forEach(async student => {
            fs.writeFile(path.join(__dirname, `Факультеты/${fac}/${student.name}.txt`), printStudent(student), function (err) {
                if (err) throw err;
            });
            await getFile(student.motivationLink.split('.')[0]);
            console.log('   ');
        })
    });
}

function printStudent(student) {
    const {name, link, phone, school, city, faculty, speciality, studyBefore, whichFaculty, motivationLink, mail} = student;
    return `Имя: ${name}
    Ссылка на соц.сеть: ${link}
    Телефон: ${phone}
    Школа: ${school}
    Город: ${city}
    Желаемый факультет: ${faculty.join(', ')}
    Специальность: ${speciality.join(', ')}
    Учился раньше: ${studyBefore}
    Где: ${whichFaculty}
    Мотивация: ${motivationLink}
    Почта: ${mail}`
}

function parseStudent(student) {
    return {
        name: student['фио'],
        link: student['ссылканааккаунтвсоциальнойсетивконтакте'],
        phone: student['мобильныйтелефон'],
        school: student['учреждениеобразованиясредняяшколагимназиялицейкласс'],
        city: student['городпроживания'],
        faculty: student['планируемыйыефакультеты'].split(','),
        speciality: student['планируемаяыеспециальностьи'].split(','),
        studyBefore: student['тыужеучаствовалвпроектестудентбгунанеделю'].toLowerCase() === 'да' ? true : false,
        whichFaculty: student['еслидатонакакомфакультетеидокакогоэтапатыпрошелотправилмотивационноеписьмовыполнилзадания2турасталстудентомбгунанеделю'],
        motivationLink: student['прикрепимотивационноеписьмонатемупочемуяхочустатьстудентомбгунанеделю'],
        mail: student['адресэлектроннойпочты'],
    }
}

accessSpreadsheet();

async function getFile(link) {
    const scopes = [
        'https://www.googleapis.com/auth/drive'
    ];

    const auth = new google.auth.JWT(
        creds.client_email, null, creds.private_key, scopes
    );

    const drive = google.drive({ version: 'v3', auth });
    let token = await auth.getAccessToken();
    // console.log(token);


    let id = link.split('=')[1]
    let test = fs.createWriteStream(`./Test/${getFileName(id)}`)

    const res = await fetch(`https://www.googleapis.com/drive/v3/files/${id}?alt=media`, {
        method: 'GET',
        headers: {
            "authorization": `Bearer ${token.token}`
        }
    });
    if (!res.ok) {
        throw new Error(`Could not fetch ${url}, received ${res.status}`);
    }

    res.body.pipe(test);

}

// getFile('https://drive.google.com/open?id=17K3aKZ_cq80g5w-aJB9jVGYchdx57fg-');


async function getFileName(id) {
    const scopes = [
        'https://www.googleapis.com/auth/drive'
    ];
    
    const auth = new google.auth.JWT(
        creds.client_email, null, creds.private_key, scopes
    );

    const drive = google.drive({ version: 'v3', auth });

    let test = fs.createWriteStream('./test.doc')
    let name = await drive.files.get({
        fileId: id,
        // mimeType: ['application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/pdf', 'application/vnd.google-apps.document']
    });
    console.log(name.data.name)
    return name.data.name;
    
    // res.pipe(test);
    // drive.files.g
    // console.log(res);
}

// getFileName('17K3aKZ_cq80g5w-aJB9jVGYchdx57fg-');