const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');
const mkdirp = require('async-mkdirp');
const path = require('path');
const rimraf = require("rimraf");
const fs = require('fs');

const url = require('url');
const { app, BrowserWindow } = require('electron');


const { google } = require('googleapis');
const fetch = require('node-fetch');

const creds = require('./secret2.json');
const faculties = new Map();

const waitFor = (ms) => new Promise(r => setTimeout(r, ms))

const asyncForEach = async (array, callback) => {
    for (let index = 0; index < array.length; index++) {
        await callback(array[index], index, array)
    }
}

const motivationLinks = [];

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
    await rimraf.sync(path.join(__dirname, 'Test'));
    await mkdirp(path.join(__dirname, `Факультеты`))
    await mkdirp(path.join(__dirname, `Test`))

    await Promise.all(Array.from(faculties).map(async faculty => {
        await mkdirp(path.join(__dirname, `Факультеты/${faculty[0]}`))

        await Promise.all(faculty[1].map(async (student, index) =>{
            await mkdirp(path.join(__dirname, `Факультеты/${faculty[0]}/${abr[faculty[0]]}-${index}`))
            fs.writeFile(path.join(__dirname, `Факультеты/${faculty[0]}/${abr[faculty[0]]}-${index}/${student.name}.txt`), printStudent(student), function (err) {
                if (err) throw err;
            });

            await Promise.all(student.motivationLink.map(async link => {
                await getFile(link.trim(), faculty[0], student.name, index);
            }))

        }))

    }))
    fs.writeFile(path.join(__dirname, `log.txt`), JSON.stringify(motivationLinks), function (err) {
        if (err) throw err;
    })
    testLinks();
}
let downloaded = 0;
function testLinks() {
    if(downloaded < motivationLinks.length){
        getMotivation(motivationLinks[downloaded])
        downloaded++;
        setTimeout(() => {
            testLinks()
        }, 500)
    }   
}


async function getMotivation(link) {
    await getMotivationFile(link.id, link.link, link.fac, link.name, link.index);
    return 'Done!';
}

function printLink(links) {
    const { id, link, fac, name, index } = links;

    return JSON.stringify(links)
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
        motivationLink: student['прикрепимотивационноеписьмонатемупочемуяхочустатьстудентомбгунанеделю'].split(','),
        mail: student['адресэлектроннойпочты'],
    }
}


const scopes = [
    'https://www.googleapis.com/auth/drive'
];

const auth = new google.auth.JWT(
    creds.client_email, null, creds.private_key, scopes
);

const drive = google.drive({ version: 'v3', auth });

async function getFile(link, fac, name, index) {
    let id = link.split('=')[1]
    motivationLinks.push({
        id: id,
        link: `https://www.googleapis.com/drive/v3/files/${id}?alt=media`,
        fac,
        name,
        index: index
    })
}

async function getMotivationFile(id, link, fac, studentName, index) {

    let token = await auth.getAccessToken();
    let name = await getFileName(id, studentName);
    if(name) {
        await waitFor(1000);
        // await mkdirp(path.join(__dirname, `Test/${fac}`))
        let test = fs.createWriteStream(`./Факультеты/${fac}/${abr[fac]}-${index}/${name}`)
        try {
            const res = await fetch(`https://www.googleapis.com/drive/v3/files/${id}?alt=media`, {
                method: 'GET',
                headers: {
                    "authorization": `Bearer ${token.token}`
                }
            });
            res.body.pipe(test);
            // console.log(downloaded / motivationLinks.length * 100, '%')
        } catch (err) {
            console.log(13, err)
        }
        
    
    }
}

async function getFileName(id, studentName) {
    let name = await drive.files.get({
        fileId: id,
    }).catch(err => console.log(err));
    if (name) {
        let fileName = name.data.name.split('.');
        return `${studentName}.${fileName[fileName.length - 1]}`;
    } else {
        return null;
    }
}

accessSpreadsheet();

const abr = {
    'Механико-математический факультет': 'ММФ',
    'Биологический факультет': 'БИО',
    'Факультет географии и геоинформатики': 'ГЕО',
    'Институт бизнеса': 'ИБ',
    'Институт теологии': 'ТЕО',
    'Исторический факультет': 'ИСТ',
    'МГЭИ им. А.Д. Сахарова': 'МГЭИ',
    'Факультет журналистики': 'ФЖ',
    'Факультет международных отношений': 'ФМО',
    'Факультет прикладной математики и информатики': 'ФПМИ',
    'Факультет радиофизики и компьютерных технологий': 'РФиКТ',
    'Факультет социокультурных коммуникаций': 'ФСК',
    'Факультет философии и социальных наук': 'ФФСН',
    'Физический факультет': 'ФИЗ',
    'Филологический факультет': 'ФИЛ',
    'Химический факультет': 'ХИМ',
    'Экономический факультет': 'ЭФ',
    'Географический факультет': 'ГЕО'
}