const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');
const mkdirp = require('mkdirp');
const path = require('path');
const rimraf = require("rimraf");
const fs = require('fs');

const url = require('url');
const { app, BrowserWindow } = require('electron');

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
    faculties.forEach((students, fac) => {
        mkdirp(path.join(__dirname, `Факультеты/${fac}`))
        students.forEach(student => {
            fs.writeFile(path.join(__dirname, `Факультеты/${fac}/${student.name}.txt`), printStudent(student), function (err) {
                if (err) throw err;
            });
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

// accessSpreadsheet();


