/**
 * Usage: node index.js src1 srcSheet1 src2 srcSheet2 out
 */
//// Core modules

//// External modules
const ExcelJS = require('exceljs');

//// Modules

// Create a function using javascript arrow
let findRow = (sheet, c, findMe) => {
    for (let r = 1; r <= 70; r++) {

        let columnValue = sheet.getCell(`${c}${r}`).value + ''
        let IDnumber = sheet.getCell(`B${r}`).value + ''
        let gender = sheet.getCell(`D${r}`).value + ''

        columnValue = columnValue.replace(/,/g, ' ') // Replace comma with space
        columnValue = columnValue.replace(/\s\s+/g, ' ') // Replace multiple spaces with 1 space

        // format
        let regex = new RegExp(`${findMe}*`, "i")
        if (columnValue.match(regex)) {
            // console.log(columnValue, '==', findMe)
            return {
                row: r,
                IDnumber: IDnumber,
                gender: gender,
            }
        }

    }
    return -1
}

    ; (async () => {
        try {
            // Process commandline args with spaces
            var args = process.argv.slice(2);
            args = args.join(' ')
            args = args.replace(/^'/, '')
            args = args.replace(/'$/, '')
            args = args.split("' '")

            let src1 = args[0]
            let srcSheet1 = args[1]
            let src2 = args[2]
            let srcSheet2 = args[3]
            let out = args[4]


            // Excel containing graduate list
            let workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(src1);
            // Select worksheet to use
            let sheet = await workbook.getWorksheet(srcSheet1)

            // Enrollment list source
            let workbookSrc = new ExcelJS.Workbook();
            await workbookSrc.xlsx.readFile(src2);
            // Select source worksheet
            let sheetSrc = await workbookSrc.getWorksheet(srcSheet2)

            // Offset header to start with actual rows
            for (let r = 5; r <= 61; r++) {

                let lastName = sheet.getCell(`C${r}`).value
                let firstName = sheet.getCell(`D${r}`).value

                let x = findRow(sheetSrc, 'C', lastName + ' ' + firstName)
                if (x != -1) {
                    sheet.getCell(`A${r}`).value = x.IDnumber
                    sheet.getCell(`F${r}`).value = x.gender
                }
            }

            // Save file
            await workbook.xlsx.writeFile(out);

        } catch (err) {
            console.log(err)
        } finally {

        }
    })()


