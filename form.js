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

async function accessSpreadsheet(id) {
    const doc = new GoogleSpreadsheet(id);
    await promisify(doc.useServiceAccountAuth)(creds);
    const info = await promisify(doc.getInfo)()
    const sheet = info.worksheets[0];

    const rows = await promisify(sheet.getRows)({
        offset: 0
    });
    rows.forEach(row => {
        let student = parseStudent(row);
        let studentName = getName(student);
        student.faculty.forEach(fac => {
            const facName = fac.trim();
            const faculty = faculties.get(facName);

            if (faculty){
                faculty.push({ ...student, indexName: studentName });
            } else {
                faculties.set(facName, [ { ...student, indexName: studentName } ])
            }
        })
    })

    await rimraf.sync(path.join(__dirname, '../../Факультеты'));
    await mkdirp(path.join(__dirname, `../../Факультеты`))

    await Promise.all(Array.from(faculties).map(async faculty => {
        await mkdirp(path.join(__dirname, `../../Факультеты/${faculty[0]}`))

        await Promise.all(faculty[1].map(async (student, index) =>{
            await mkdirp(path.join(__dirname, `../../Факультеты/${faculty[0]}/`))
            try{
                fs.writeFile(path.join(__dirname, `../../Факультеты/${faculty[0]}/${student.indexName}.txt`), printStudent(student), function (err) {
                    if (err) throw err;
                });
            } catch(err){
                document.getElementById('error').innerHTML = err;
            }

            await Promise.all(student.motivationLink.map(async (link, indexLink) => {
                await getFile(link.trim(), faculty[0], student.name, index, student.indexName, indexLink);
            }))

        }))

    }))
    // fs.writeFile(path.join(__dirname, `log.txt`), JSON.stringify(motivationLinks), function (err) {
    //     if (err) throw err;
    // })
    testLinks();
}

function getName(student) {
    let name = '';
    student.faculty.forEach(fac => {
        let facName = fac.trim();
        let abrName = abr[facName];
        const faculty = faculties.get(facName);

        if (faculty) {
            name += `${abrName}-${faculty.length + 1} `
        } else {
            name += `${abrName}-${1} `
        }
    });

    console.log(name.trim());

    return name.trim();
}

let downloaded = 0;

function testLinks() {
    if(downloaded < motivationLinks.length){
        getMotivation(motivationLinks[downloaded])
        downloaded++;
        setTimeout(() => {
            testLinks()
        }, 300)
    }   
    let progress = document.getElementById('progress');
    let percent = downloaded / motivationLinks.length * 100
    progress.setAttribute('aria-valuenow', `${ percent }`);
    progress.style.width = `${ percent }%`;
    console.log(`${ downloaded / motivationLinks.length * 100 } %`)
    if(percent === 100) {
        console.log('finish');
        document.getElementById('done').classList.remove('hidden');
        document.getElementById('doneFolder').innerHTML = `Путь к папкам факультетов: ${path.join(__dirname, '../../Факультеты')}`;
        document.getElementById('error').innerHTML = '';
        document.getElementById('startButton').disabled = false;
    }
}


async function getMotivation(link) {
    await getMotivationFile(link.id, link.link, link.fac, link.name, link.index, link.indexName, link.indexLink);
    return 'Done!';
}

function printLink(links) {
    const { id, link, fac, name, index } = links;

    return JSON.stringify(links)
}

function printStudent(student) {
    const { name, link, phone, school, city, faculty, speciality, studyBefore, whichFaculty, motivationLink, mail } = student;
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
        link: student['ссылканааккаунтвсоциальнойсетивконтакте'] || student['ссылканааккаунтввконтакте'],
        phone: student['мобильныйтелефон'],
        school: student['учреждениеобразованиясредняяшколагимназиялицейкласс'] || student['учреждениеобразованиясредняяшколагимназиялицейссузкласс'],
        city: student['городпроживания'],
        faculty: student['планируемыйыефакультеты'].split(','),
        speciality: student['планируемаяыеспециальностьи'].split(','),
        studyBefore: student['тыужеучаствовалвпроектестудентбгунанеделю'],
        whichFaculty: student['еслидатонакакомфакультетеидокакогоэтапатыпрошелотправилмотивационноеписьмовыполнилзадания2турасталстудентомбгунанеделю'] || student['еслидатонакакомфакультетеидокакогоэтапатыпрошел'],
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

async function getFile(link, fac, name, index, indexName, indexLink) {
    let id = link.split('=')[1]
    motivationLinks.push({
        id: id,
        link: `https://www.googleapis.com/drive/v3/files/${id}?alt=media`,
        fac,
        name,
        index: index,
        indexName: indexName,
        indexLink
    })
}

async function getMotivationFile(id, link, fac, studentName, index, indexName, indexLink) {
    let token = await auth.getAccessToken();
    let name = await getFileName(id, `${indexName}-${indexLink + 1}`);
    if(name) {
        await waitFor(1000);
        // await mkdirp(path.join(__dirname, `Test/${fac}`))
        let test = fs.createWriteStream(path.join(__dirname, `../../Факультеты/${fac}/${name}`))
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
        } catch (err) {
            document.getElementById('error').innerHTML = err;
        }
    } else {
        console.log('err');
    }

}

// accessSpreadsheet();

const abr = {
    'Механико-математический факультет': 'ММФ',
    'Биологический факультет': 'БИО',
    'Факультет географии и геоинформатики': 'ГЕО',
    'Институт бизнеса': 'ИБ',
    'Институт теологии': 'ТЕО',
    'Исторический факультет': 'ИСТ',
    'МГЭИ им. А.Д. Сахарова': 'МГЭИ',
    'МГЭИ им. А. Д. Сахарова': 'МГЭИ',
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