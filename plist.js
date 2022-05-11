/**
 * Promotional List
 * Usage: node plist.js
 */
//// Core modules
const fs = require('fs')
const { join } = require('path')

//// External modules
const ExcelJS = require('exceljs')

//// Modules
const kisha = require('./kisha');


; (async () => {
    try {

        let dirSrc = "H:/Enrollment and Promotional List/List Final ( Enrollment)/"
        // let dirSrc = "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2017-2018/1ST SEM"

        function getFiles(dir, files_) {
            files_ = files_ || [];
            var files = fs.readdirSync(dir).filter(item => !(/^(\..*)|^(\~.*)/g).test(item))
            for (var i in files) {
                var name = join(dir, files[i]).replace(/\\/g, '/')
                if (fs.statSync(name).isDirectory()) {
                    getFiles(name, files_);
                } else {
                    files_.push(name);
                }
            }
            return files_;
        }

        let sources = [
            // 'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2015-2016 (1st Semester).xlsx',
            // 'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2015-2016 (2nd Semester).xlsx',
            // 'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2016-2017 (1st Semester).xlsx',
            // 'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2016-2017 (2nd Semester).xlsx',
            // 'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2017-2018 (1st Semester).xlsx',
            // 'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2017-2018 (2nd Semester).xlsx',
            'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2018-2019 (1st Semester).xlsx',
            'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2018-2019 (2nd Semester).xlsx',
            'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2019-2020 (1st Semester).xlsx',
            'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2019-2020 (2nd Semester).xlsx',
            'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2020-2021 (1st Semester).xlsx',
            'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2020-2021 (2nd Semester).xlsx',
            'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2021-2022 (1st Semester).xlsx',
            'H:/Enrollment and Promotional List/List Final ( Enrollment)/Enrollment List A.Y 2021-2022 (2nd Semester).xlsx'
        ]
        let destins = [
            // 'H:/Enrollment and Promotional List/Promotional List PDF & Excel/2015 CHED Without Grades/1st sem',
            // 'H:/Enrollment and Promotional List/Promotional List PDF & Excel/2015 CHED Without Grades/2ND SEM',
            // "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2016-2017/1ST SEM",
            // "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2016-2017/2ND SEM",
            // "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2017-2018/1ST SEM",
            // "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2017-2018/2ND SEM",
            "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2018-2019/1ST SEM",
            "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2018-2019/2ND SEM",
            "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2019-2020/1ST SEM",
            "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2019-2020/2ND SEM",
            "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2020-2021/1ST SEM",
            "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2020-2021/2ND SEM",
            "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2021-2022",
            "H:/Enrollment and Promotional List/Promotional List PDF & Excel/2021-2022",
        ]

        let files = []
        sources.forEach((source, sourceIndex) => {
            let name = source
            let files2 = getFiles(destins[sourceIndex]).filter(item => (/(.*\.xlsx)$/g).test(item)).filter(item => !(/(.*mos.*)/g).test(item))
            files2.forEach((name2) => {
                files.push({
                    name: name,
                    name2: name2
                })
            })
        })
        

        let saveds = []
        let ignoreds = []

        for (let f = 0; f < files.length; f++) {
            let file = files[f]

            let book1 = await kisha.getExcel(file.name)
            let sheet = book1.worksheets[0]
            if (book1.worksheets.length > 1) {
                throw new Error('More than 1 unused source sheet.')
            }

            let book2 = new ExcelJS.Workbook();
            await book2.xlsx.readFile(file.name2);

            for (let b = 0; b < book2.worksheets.length; b++) {
                let sheet2 = book2.worksheets[b]
                let sheet2name = sheet2.name
                let workbookTpl = new ExcelJS.Workbook();
                await workbookTpl.xlsx.readFile(`D:/nodejs/excel-scripts/templates/promo-list.xlsx`);
                let sheetTpl = await workbookTpl.getWorksheet(`Sheet1`)


                let nextDestRow = 0
                let totalRows = kisha.getWorksheetTotalRows(sheet2)
                sheet.eachRow(function (rowSrc, rowSrcIndex) {
                    // Work on rows with ID only
                    let criteria = new RegExp(`^GSC-*`)

                    if (typeof rowSrc.getCell(`B`).value === 'string' && rowSrc.getCell(`B`).value.match(criteria) && typeof rowSrc.getCell(`E`).value === 'string' && rowSrc.getCell(`E`).value.trim().includes('MOS')) {
                        let gscID = rowSrc.getCell(`B`).value.trim()
                        // console.log(gscID)

                        let startRow = -1

                        for (let rowIndex = 1; rowIndex <= totalRows; rowIndex++) {
                            if (sheet2.getCell(`B${rowIndex}`).value && sheet2.getCell(`B${rowIndex}`).value.match(criteria)) {
                                if (startRow >= 1) {
                                    let endRow = rowIndex - 2 // TODO: - 2 might not work if there are more spaces
                                    let rowSpan = endRow - startRow

                                    let letters = 'ABCDEFGHIJKLMNOP'.split('')
                                    letters.forEach((letter) => {
                                        sheetTpl.getCell(`${letter}${nextDestRow}`).value = sheet2.getCell(`${letter}${startRow}`).value
                                    })
                                    // grades
                                    for (let y = 1; y <= rowSpan; y++) {
                                        let offsetSrc = startRow + y
                                        let offsetDest = nextDestRow + y
                                        'HIJKLMNOP'.split('').forEach((letter) => {
                                            sheetTpl.getCell(`${letter}${offsetDest}`).value = sheet2.getCell(`${letter}${offsetSrc}`).value
                                        })
                                    }
                                    nextDestRow += rowSpan + 1

                                    break
                                }
                                // FOUND!
                                if (sheet2.getCell(`B${rowIndex}`).value.trim() === gscID) {
                                    nextDestRow++
                                    startRow = rowIndex

                                }
                            }
                        }


                    }
                })



                if (nextDestRow <= 0) {
                    ignoreds.push(file.name2)

                } else {
                    let suffix = sheet2name.toLowerCase().replace(/\s+/g, '_') // Replace multiple spaces
                    let out = `${file.name2.replace('.xlsx', '')}-mos-${suffix}.xlsx`
                    saveds.push(out)
                    await workbookTpl.xlsx.writeFile(out)
                }


            }


        }
        console.log(`Saved ${saveds.length} files:`, saveds)
        console.log(`Ignored ${ignoreds.length} files:`, ignoreds)
    } catch (err) {
        console.log(err)
    } finally {

    }
})()


