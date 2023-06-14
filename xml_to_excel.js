const XlsxPopulate = require('xlsx-populate');
const fs = require('fs');
const xml2js = require('xml2js');

// Function for parsing the input XML file
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
    inputData.then(data => {
        const periodsArr = data.timetable.periods;

        const daysArr = data.timetable.daysdefs;

        // // Iterate over each faculty
        for (let facultyIndex = 1; facultyIndex <= 35; facultyIndex++) {
            const facultyName = `Faculty ${facultyIndex}`;
            const facultyWorksheet = workbook.addSheet(facultyName);

            for (let classIndex = 1; classIndex <= 15; classIndex++) {
                const className = `Class ${classIndex}`;
                const classWorksheet = workbook.addSheet(`${facultyName} - ${className}`);

                var periodsObj = periodsArr[0];
                var periods = periodsObj.period;

                var daysObj = daysArr[0];
                var days = daysObj.daysdef;

                var rowNumber = 1;
                var heightInPixels = 30;
                var widthInPixels = 20;

                const row = classWorksheet.row(rowNumber);
                row.height(heightInPixels);
                
                for (let i = 0; i < periods.length; i++) {
                    const period = periods[i];
                    const startTime = period.$.starttime;
                    const endTime = period.$.endtime;
                    const startTimeSlot = formatTimeSlot(startTime);
                    const endTimeSlot = formatTimeSlot(endTime);
                    const timeSlot = `${startTimeSlot} to \n ${endTimeSlot}`;

                    const cell = classWorksheet.cell(1, i + 2);
                    cell.value(timeSlot);
                    cell.style({
                        fill: '00FF00',
                        bold: true,
                        fontSize: 13,
                        horizontalAlignment: 'center',
                        verticalAlignment: 'center',
                        wrapText: true
                    });
                    const column = classWorksheet.column(i + 2);
                    column.width(widthInPixels);
                }

                var list_of_days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Sunday'];
                for (let i = 0; i < days.length; i++) {
                    const day = days[i];
                    const day_name = day.$.name;

                    const height_pxls = 50;
                    
                    const col = classWorksheet.column(1);
                    col.width(widthInPixels);
                  
                    if(list_of_days.includes(day_name)){
                        const cell = classWorksheet.cell(i, 1);
                        cell.value(day_name);
                        cell.style({
                            fill: '00FF00',
                            bold: true,
                            fontSize: 13,
                            horizontalAlignment: 'center',
                            verticalAlignment: 'center'
                        });

                        const rw = classWorksheet.row(i);
                        rw.height(height_pxls);
                    }
                }
            }
        }

        workbook.toFileAsync('Output_Excel File.xlsx').then(() => {
            console.log('Output file saved successfully!');
        }).catch((error) => {
            console.error('An error occurred while saving the output file:', error);
        });
    }).catch(error => {
        console.error('An error occurred while parsing the XML:', error);
    });
});

function formatTimeSlot(time){
    const hour = parseInt(time.split(':')[0]);
    let formattedTime = '';
    
    if (hour >= 12) {
        formattedTime = `${hour % 12 === 0 ? 12 : hour}:${time.split(':')[1]} PM`;
    }
    else {
        formattedTime = `${hour === 0 ? 12 : hour}:${time.split(':')[1]} AM`;
    }
    
    return formattedTime;
}
