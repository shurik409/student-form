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
    mkdirp(path.join(__dirname, `Факультеты`))
    mkdirp(path.join(__dirname, `Test`))
    // faculties.forEach((students, fac) => {
    //     mkdirp(path.join(__dirname, `Факультеты/${fac}`))
    //     students.forEach(async student => {
    //         fs.writeFile(path.join(__dirname, `Факультеты/${fac}/${student.name}.txt`), printStudent(student), function (err) {
    //             if (err) throw err;
    //         });
    //         student.motivationLink.forEach(link => {
    //             getFile(link.trim());
    //             // await wait(500);
    //         })

    //         // asyncForEach(student.motivationLink, async link => {
    //         //     await getFile(link.trim());
    //         //     await wait(500);
    //         // })
    //     })
    // });

    await Promise.all(Array.from(faculties).map(async faculty => {
        mkdirp(path.join(__dirname, `Факультеты/${faculty[0]}`))

        await Promise.all(faculty[1].map(async student =>{
            fs.writeFile(path.join(__dirname, `Факультеты/${faculty[0]}/${student.name}.txt`), printStudent(student), function (err) {
                if (err) throw err;
            });

            await Promise.all(student.motivationLink.map(async link => {
                await getFile(link.trim(), faculty[0], student.name);
                // await wait(500);
            }))

        }))

        // faculty[1].forEach(async student => {
        //     fs.writeFile(path.join(__dirname, `Факультеты/${faculty[0]}/${student.name}.txt`), printStudent(student), function (err) {
        //         if (err) throw err;
        //     });
        //     // student.motivationLink.forEach(link => {
        //     //     getFile(link.trim());
        //     //     // await wait(500);
        //     // })
        // })
    }))
    // motivationLinks.map(link => console.log(link));
    // let num = 0;
    // await Promise.all(motivationLinks.map(async link => {
    //     const file = await getMotivation(link);
    //     // console.log(file);
    //     // console.log(num);
    //     // num++;
    // })).catch(err => console.log(12, err))
    // console.log('Done2!');
    testLinks();
}
let downloaded = 0;
function testLinks() {
    if(downloaded < motivationLinks.length){
        getMotivation(motivationLinks[downloaded])
        downloaded++;
        console.log(1);
        setTimeout(() => {
            testLinks()
        }, 500)
    }   
}


async function getMotivation(link) {
    await getMotivationFile(link.id, link.link, link.fac, link.name);
    return 'Done!';
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

accessSpreadsheet();

const scopes = [
    'https://www.googleapis.com/auth/drive'
];

const auth = new google.auth.JWT(
    creds.client_email, null, creds.private_key, scopes
);

const drive = google.drive({ version: 'v3', auth });

async function getFile(link, fac, name) {
    
    // console.log(token);


    let id = link.split('=')[1]
    // console.log(id);
    // let name = await getFileName(id);
    motivationLinks.push({
        id: id,
        link: `https://www.googleapis.com/drive/v3/files/${id}?alt=media`,
        fac,
        name
    })
    // console.log(motivationLinks);
    // let test = fs.createWriteStream(`./Test/${name}`)

    // const res = await fetch(`https://www.googleapis.com/drive/v3/files/${id}?alt=media`, {
    //     method: 'GET',
    //     headers: {
    //         "authorization": `Bearer ${token.token}`
    //     }
    // });
    // if (!res.ok) {
    //     throw new Error(`Could not fetch ${`https://www.googleapis.com/drive/v3/files/${id}?alt=media`}, received ${res.status}`);
    // }

    // res.body.pipe(test);

}

async function getMotivationFile(id, link, fac, studentName) {

    let token = await auth.getAccessToken();
    let name = await getFileName(id, studentName);
    if(name) {
        await waitFor(1000);
        mkdirp(path.join(__dirname, `Test/${fac}`))
        let test = fs.createWriteStream(`./Test/${fac}/${name}`)
        try {
            const res = await fetch(`https://www.googleapis.com/drive/v3/files/${id}?alt=media`, {
                method: 'GET',
                headers: {
                    "authorization": `Bearer ${token.token}`
                }
            });
            res.body.pipe(test);
        } catch (err) {
            console.log(13, err)
        }
        
    
    }
}

// getFile('https://drive.google.com/open?id=17K3aKZ_cq80g5w-aJB9jVGYchdx57fg-');


async function getFileName(id, studentName) {
    // let test = fs.createWriteStream('./test.doc')
    let name = await drive.files.get({
        fileId: id,
        // mimeType: ['application/msword', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/pdf', 'application/vnd.google-apps.document']
    }).catch(err => console.log(err));
    // console.log(name.data.name)
    if (name) {
        let fileName = name.data.name.split('.');
        return `${studentName}.${fileName[fileName.length - 1]}`;
    } else {
        return null;
    }
    
    // res.pipe(test);
    // drive.files.g
    // console.log(res);
}

// getFileName('17K3aKZ_cq80g5w-aJB9jVGYchdx57fg-');