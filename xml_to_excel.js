const XlsxPopulate = require('xlsx-populate');
const fs = require('fs');
const xml2js = require('xml2js');
const { error } = require('console');

// Read the input XML file
function parseXML(filePath){
    const xmlData = fs.readFileSync(filePath, 'utf-8');

    return new Promise((resolve, reject) => {
        const parser = new xml2js.Parser();
        parser.parseString(xmlData, (error, result) => {
            if(error){
                reject(error);
            }
            else{
                resolve(result);
            }
        });
    });
}

const inputData = parseXML('inputxmlFile.xml');
console.log(inputData);

// Create a new Excel workbook
XlsxPopulate.fromBlankAsync().then(workbook => {
    // const worksheet = workbook.addSheet('Sheet-1');
    // worksheet.cell('A1').value('Hello, World!');

    // inputData.then(data => {
    //     const periodsArr = data.timetable.periods;
    //     // console.log('Periods Array: ',periodsArr);
    //     // const periodsString = JSON.stringify(periodsArr, null, 2);
    //     // console.log(periodsString);
    //     console.log(periodsArr);

    //     // // Iterate over each faculty
    //     for (let facultyIndex = 1; facultyIndex <= 35; facultyIndex++) {
    //         const facultyName = `Faculty ${facultyIndex}`;
    //         const facultyWorksheet = workbook.addSheet(facultyName);

    //         for (let classIndex = 1; classIndex <= 15; classIndex++) {
    //             const className = `Class ${classIndex}`;
    //             const classWorksheet = workbook.addSheet(`${facultyName} - ${className}`);

    //             var periodsObj = periodsArr[0];
    //             var periods = periodsObj.period;
                
    //             for (let i = 0; i < periods.length; i++) {
    //                 const period = periods[i];
    //                 const startTime = period.$.starttime;
    //                 const endTime = period.$.endtime;
    //                 const timeSlot = `${startTime} to ${endTime}`;
                  
    //                 // Set the time value in the first row, starting from column B
    //                 const cell = classWorksheet.cell(1, i + 2); // Assuming "facultyWorksheet" is the worksheet object for the faculty
    //                 cell.value(timeSlot);
    //             }
    //         }
    //     }

    //     workbook.toFileAsync('Output_Excel File.xlsx').then(() => {
    //         console.log('Output file saved successfully!');
    //     }).catch((error) => {
    //         console.error('An error occurred while saving the output file:', error);
    //     });
    // }).catch(error => {
    //     console.error('An error occurred while parsing the XML:', error);
    // });
});

// // // Iterate over each faculty
// for (let facultyIndex = 1; facultyIndex <= 35; facultyIndex++) {
//   const facultyName = `Faculty ${facultyIndex}`;

//   // Create a new worksheet for the faculty
//   const facultyWorksheet = workbook.addSheet(facultyName);

//   // Iterate over each class
//   for (let classIndex = 1; classIndex <= 15; classIndex++) {
//     const className = `Class ${classIndex}`;

//     // Create a new worksheet for the class under the faculty
//     const classWorksheet = workbook.addSheet(`${facultyName} - ${className}`);

//     // Get the timetable data for the current faculty and class
//     const timetableData = getTimetableData(inputData, facultyName, className);

//     // Write the timetable data to the class worksheet
//     classWorksheet.cell('A1').value(timetableData);
//   }
// }

// // Save the workbook to the output Excel file
// workbook.toFileAsync('Output_Excel File.xlsx')
//   .then(() => {
//     console.log('Output file saved successfully!');
//   })
//   .catch((error) => {
//     console.error('An error occurred while saving the output file:', error);
//   });
