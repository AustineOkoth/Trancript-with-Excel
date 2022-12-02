const express = require('express'); // initializing the main express module
const bodyParser = require("body-parser");
let xlsx = require("xlsx") //xlsx module that enables that makes it possible read Exel data
let workBook = xlsx.readFile("./views/spreadsheet.xlsx");
dataToJson = xlsx.utils.sheet_to_json(workBook.Sheets["Sheet1"])

let data = []; //Individuals marks in each subject
let individualSubjectGrade = []; //Individual Grades ie A,B,C....

let totalNumerical = []; //All students marks in numerical Analysis
let totalReal = []; //All students marks in Real Analysis
let totalComps = []; //All students marks in Computer Graphics
let totalODE = []; //All students marks in ODE
let totalScientific = []; //All students marks in Scientific
let totalAccounts = []; //All students marks in Acoounts
let totalOS = [] //All students marks in Operating System


function generateAllMarks() {
    Object.keys(dataToJson).forEach((key) => {
        let numkey = dataToJson[key]
        totalNumerical.push(numkey.NUMERICAL)
        totalReal.push(numkey['REAL ANAYSIS']);
        totalComps.push(numkey['COMPUTER GRAPHICS'])
        totalODE.push(numkey.ODE)
        totalScientific.push(numkey.SCIENTIFIC)
        totalAccounts.push(numkey['ACCOUNTS & FINANCE'])
        totalOS.push(numkey['OPERATING SYSTEM'])
    })
}
generateAllMarks() // FUNCTION TO GENERATE TOTAL MARKS IN ALL UNITS

let sum1 = 0; let sum2 = 0; let sum3 = 0; let sum4 = 0; let sum5 = 0; let sum6 = 0; let sum7 = 0;
for (i = 0; i < totalNumerical.length; i++) {
    //LOOP TO GET THE AVERAGE OF EACH SUBJECT
    sum1 += totalNumerical[i];
    sum2 += totalReal[i];
    sum3 += totalComps[i];
    sum4 += totalODE[i];
    sum5 += totalScientific[i];
    sum6 += totalAccounts[i];
    sum7 += totalOS[i];

    average1 = Math.ceil(sum1 / totalNumerical.length);
    average2 = Math.ceil(sum2 / totalNumerical.length);
    average3 = Math.ceil(sum3 / totalNumerical.length);
    average4 = Math.ceil(sum4 / totalNumerical.length);
    average5 = Math.ceil(sum5 / totalNumerical.length);
    average6 = Math.ceil(sum6 / totalNumerical.length);
    average7 = Math.ceil(sum7 / totalNumerical.length);
}
const app = express();
app.use(express.static('./views'));
app.use(express.json());
app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));

app.get("/", (req, res) => {
    res.status(200).render("registration")
})

app.post("/result", (req, res) => {
    const { regNumber } = req.body //Acquiring the Registration Number by the user in the first page 
    dataToJson.forEach((result) => {
        Object.keys(result).forEach(async (key) => {
            let i = result[key]
            if (i == `${regNumber}`) {
                data.push((result.ADDMISION));
                data.push((result.NAMES));
                data.push(result.NUMERICAL);
                data.push(result['REAL ANAYSIS']);
                data.push(result['COMPUTER GRAPHICS']);
                data.push(result.ODE);
                data.push(result.SCIENTIFIC);
                data.push(result['ACCOUNTS & FINANCE']);
                data.push(result['OPERATING SYSTEM']);
            }
        });
    });

    let totalSum = data[2] + data[3] + data[4] + data[5] + data[6] + data[7] + data[8];
    let averageScore = Math.ceil(totalSum / 7);

    for (i = 0; i < data.length; i++) {
        //FOR LOOP TO CHECK THE MARKS OF A PERSON AND GIVE THEM APPROPRIATE GRADES
        if (data[i] >= 70) {
            let grade = individualSubjectGrade.push("A");
        }
        if (data[i] >= 60 && data[i] < 70) {
            let grade = individualSubjectGrade.push("B");
        }
        if (data[i] >= 50 && data[i] < 60) {
            let grade = individualSubjectGrade.push("C");
        }
        if (data[i] >= 40 && data[i] < 50) {
            let grade = individualSubjectGrade.push("D");
        }
        if (data[i] < 40) {
            let grade = individualSubjectGrade.push("E");
        }
    }
    res.status(200).render("index", {
        //Dynamically replacing the variables 
        regNumber: `${data[0]}`, nameOfStudent: `${data[1]}`, marksNumerical: `${data[2]}`,
        GRADENUMERICAL: `${individualSubjectGrade[0]}`, marksReal: `${data[3]}`, GRADEREAL: `${individualSubjectGrade[1]}`,
        marksComps: `${data[4]}`, GRADECOMPS: `${individualSubjectGrade[2]}`, marksODE: `${data[5]}`,
         GRADEODE: `${individualSubjectGrade[3]}`, marksScientific: `${data[6]}`, 
         GRADECOMPUTING: `${individualSubjectGrade[4]}`, marksAccounts: `${data[7]}`, 
         GRADEACCOUNTS: `${individualSubjectGrade[5]}`,
        marksOS: `${data[8]}`, GRADEOS: `${individualSubjectGrade[6]}`, totalMarks: `${totalSum}`,
        averageScore: `${averageScore}`, avgNum: `${average1}`, avgReal: `${average2}`, avgComp: `${average3}`,
         avgODE: `${average4}`, avgScientific: `${average5}`, avgAccounts: `${average6}`, avgOS: `${average7}`
    });
})
app.listen("5000", () => {
    console.log("Listening on port 5000");
})