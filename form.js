const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');
// const mkdirp = require('mkdirp');
const mkdirp = require('async-mkdirp');
const path = require('path');
const rimraf = require("rimraf");
const fs = require('fs');

const url = require('url');
const { app, BrowserWindow } = require('electron');

const creds = require('./secret2.json');
const faculties = new Map();

async function accessSpreadsheet(id) {
    const doc = new GoogleSpreadsheet(id);
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
    await rimraf.sync(path.join(__dirname, '../../Факультеты'));
    await mkdirp(path.join(__dirname, `../../Факультеты`))
    faculties.forEach(async (students, fac) => {
        await mkdirp(path.join(__dirname, `../../Факультеты/${fac}`))
        students.forEach(student => {
            try{
                fs.writeFile(path.join(__dirname, `../../Факультеты/${fac}/${student.name}.txt`), printStudent(student), function (err) {
                    if (err) {
                        throw err;
                    }
                }); 
            } catch(err){
                document.getElementById('error').innerHTML = err;
            }
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

async function addFacultyFolder(obj) {
    obj.disabled = true;
    let link = document.getElementById('link').value.split('/');
    let id = '';
    link.forEach((value, index) => {
        if (value === 'd') {
            id = link[index + 1];
        }
    })
    if (id.length) {
        try {
            await accessSpreadsheet(id);
            console.log('finish');
            document.getElementById('done').classList.remove('hidden');
            document.getElementById('doneFolder').innerHTML = `Путь к папкам факультетов: ${path.join(__dirname, '../../Факультеты')}`;
            document.getElementById('error').innerHTML = '';
        } catch (err) {
            document.getElementById('error').innerHTML = err;
        }
    } else {
        console.log('err');
    }
    obj.disabled = false;

}

// accessSpreadsheet();


