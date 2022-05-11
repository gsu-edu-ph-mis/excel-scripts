const ExcelJS = require('exceljs')

class Kisha {
    workbook
    sheet
    cell
    constructor(workbook, sheetName) {
        this.workbook = workbook;
        this.setSheet(sheetName)
    }
    setSheet(sheetName) {
        this.sheet = this.workbook.getWorksheet(sheetName)
        return this
    }
    getRows() {
        const rows = this.sheet.getColumn(1);
        const rowsCount = rows['_worksheet']['_rows'].length;
        for (let r = 1; r <= rowsCount; r++) {

        }
    }
    mergeCells(range) {
        this.sheet.mergeCells(range)
        let cells = range.split(':')
        this.getCell(cells[0])
        return this
    }
    setCell(cell) {
        this.cell = cell
        return this
    }
    getCell(cell) {
        this.cell = this.sheet.getCell(cell)
        return this
    }
    value(s) {
        this.cell.value = s
        return this
    }
    numFmt(s) {
        this.cell.numFmt = s
        return this
    }
    align(pos) {
        if (['top', 'middle', 'bottom'].includes(pos)) {
            lodash.set(this, 'cell.alignment.vertical', pos)
        }
        if (['left', 'center', 'right'].includes(pos)) {
            lodash.set(this, 'cell.alignment.horizontal', pos)
        }
        return this
    }
    wrapText(s) {
        lodash.set(this, 'cell.alignment.wrapText', s)
        return this
    }
    font(s) {
        lodash.set(this, 'cell.font.name', s)
        return this
    }
    fontSize(s) {
        lodash.set(this, 'cell.font.size', s)
        return this
    }
    fontColor(s) {
        lodash.set(this, 'cell.font.color.argb', s)
        return this
    }
    bold(s) {
        lodash.set(this, 'cell.font.bold', s)
        return this
    }
    italic(s) {
        lodash.set(this, 'cell.font.italic', s)
        return this
    }
    underline(s) {
        lodash.set(this, 'cell.font.underline', s)
        return this
    }
    bgFill(s) {
        lodash.set(this, 'cell.fill', {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: s }
        })

        return this
    }
    border(t, r, b, l) {
        if (t && r === undefined && b === undefined && l === undefined) {
            lodash.set(this, 'cell.border.top.style', t)
            lodash.set(this, 'cell.border.right.style', t)
            lodash.set(this, 'cell.border.bottom.style', t)
            lodash.set(this, 'cell.border.left.style', t)
            return this
        }
        if (t) {
            lodash.set(this, 'cell.border.top.style', t)
        }
        if (r) {
            lodash.set(this, 'cell.border.right.style', r)
        }
        if (b) {
            lodash.set(this, 'cell.border.bottom.style', b)
        }
        if (l) {
            lodash.set(this, 'cell.border.left.style', l)
        }
        return this
    }
}


module.exports = {
    create: async (file, sheetName = 'Sheet1') => {
        let workbook = new ExcelJS.Workbook()
        await workbook.xlsx.readFile(file)
        return new Kisha(workbook, sheetName)
    },
    getExcel: async (file) => {
        let workbook = new ExcelJS.Workbook();
        return workbook.xlsx.readFile(file);
    },
    getWorksheet: async (file, sheetName = 'Sheet1') => {
        let workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(file);
        return workbook.getWorksheet(sheetName)
    },
    getWorksheetTotalRows: (sheet) => {
        return sheet.getColumn(1)['_worksheet']['_rows'].length
    }
}