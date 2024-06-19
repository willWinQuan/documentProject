
/**
 * write by Chen Hai Quan --2020/05/06
 * V1.end by Chen Hai Quan ---2020/05/29
 */
/**
 * 1.修復margeCells border -chen hai quan --2020/08/07
 */

const Excel = require('exceljs');

//公共類 --------begin
class commonUntil {

    static numToAlphabet(num) {
        return String.fromCharCode(64 + parseInt(num));
    }

    static getStyleObj(Style, styleClass) {
        let styleObj = {};
        for (let className of styleClass) {
            Object.assign(styleObj, Style[className])
        }
        return styleObj
    }

    static getDataType(data) {
        let type = Object.prototype.toString.call(data);
        type = type.replace('[object', '').replace(']', '').trim();
        return type.toLowerCase()
    }

}
//公共類 --------end

class handlerSheet {
    constructor(WorkBook, sheet, Style) {

        this.WorkBook = WorkBook;
        this.sheet = sheet;
        this.Style = Style;

        this.sheetData = this.sheet.sheetData ? this.sheet.sheetData : [];
        this.dataRowIndex = -1;
        this.rowIndex = this.sheet.startRowIndex ? this.sheet.startRowIndex : 1;
        this.cellIndex = this.sheet.startCellIndex ? this.sheet.startCellIndex : 1;

        this.WorkSheet;
        this.row;
        this.rowType = "";
        this.addRowNum = 0;

        this.stringRowNum = -1;
        this.columnCustomPro = {
            dataSameMerge: {}
        }
        this.dataSameMergeParams = {}
    }
    async initSheet({
        SheetName = "sheet_" + Date.now(),
        Properties = {},
        Views = [],
        State = "visible",
        PageSetup = {},
        Style = { testClass: {} },
        Columns = []
    }) {
        if (Columns.length === 0) {
            throw "表格Columns為必須項,不能為空。"
        }
        this.WorkSheet = await this.WorkBook.addWorksheet(SheetName, {
            properties: Properties,
            views: Views,
            pageSetup: PageSetup
        })
        this.WorkSheet.state = State;

        Object.assign(this.Style, Style)

        this.setSheetColumns(Columns)
    }
    setSheetColumns(Columns) {
        for (let i = 0; i < Columns.length; i++) {
            if (commonUntil.getDataType(Columns[i]) === 'object') {
                this.WorkSheet.getColumn(this.cellIndex + i).key = Columns[i].key;
                this.WorkSheet.getColumn(this.cellIndex + i).width = Columns[i].width;
                this.columnCustomPro.dataSameMerge[Columns[i].key] = Columns[i].dataSameMerge ? Columns[i].dataSameMerge : false;
            } else {
                this.WorkSheet.getColumn(this.cellIndex + i).key = Columns[i]
            }
        }
    }
    async processor() {
        if (this.sheetData.length === 0) {
            throw "sheetData未设置"
        }
        for (let i = 0; i < this.sheetData.length; i++) {
            this.dataRowIndex = i;
            switch (this.sheetData[i].type) {
                case "stringRow":
                    await this.handlerStringRow(this.sheetData[i]);
                    break;
                default:
                    await this.handlerStringRow(this.sheetData[i])
            }
        }
    }
    async handlerStringRow({
        rowConfig = {
            height: 15,
            hidden: false,
            outlineLevel: 0,
            values: [],
            type: ""
        },
        styleClass = [],
        styleCol = {},
        values = []
    }) {
        await this.setRowConfig(rowConfig);
        let styleObj = {};
        Object.assign(styleObj, commonUntil.getStyleObj(this.Style, styleClass), styleCol)
        await this.setStringRowValues(values, styleObj);
    }

    setRowConfig(rowConfig) {
        this.row = this.WorkSheet.getRow(this.rowIndex + this.addRowNum);
        this.rowType = rowConfig.type;
        delete rowConfig.type;
        Object.assign(this.row, rowConfig);
    }
    async setStringRowValues(values, styleObj) {
        for (let i = 0; i < values.length; i++) {
            let style = {}, cell;

            if (commonUntil.getDataType(values[i]) === "object") {
                Object.assign(style,
                    styleObj,
                    (values[i].styleClass
                        ? commonUntil.getStyleObj(this.Style, values[i].styleClass)
                        : {}
                    ),
                    (values[i].styleCol ? values[i].styleCol : {})
                );

                if (commonUntil.getDataType(values[i].value) === "array") {
                    for (let j = 0; j < values[i].value.length; j++) {
                        await this.handlerStringRow(values[i].value[j]);
                        if (j === values[i].value.length - 1) {

                            let beginRow = this.WorkSheet.getRow(this.row._number - j);
                            let ArrayBeginCellKey = values[i].value[j].values[0].key;
                            let endCellIndex = this.row.getCell(ArrayBeginCellKey)._column._number;

                            for (let cellNum = this.cellIndex; cellNum < endCellIndex; cellNum++) {
                                let beginMergeAddress = beginRow.getCell(cellNum).address;
                                //合併多行的統一設置文字居中
                                beginRow.getCell(cellNum).alignment = {
                                    horizontal: "center",
                                    vertical: "middle"
                                }
                                let endMergeAddress = this.row.getCell(cellNum).address;
                                if(style.border){
                                    beginRow.getCell(cellNum).border=style.border
                                }
                                this.WorkSheet.mergeCells(beginMergeAddress, endMergeAddress)
                            }
                        }
                    }
                    this.addRowNum -= (values[i].value.length - 2) > 0 ? (values[i].value.length - 2) : 1;
                } else {
                    let key = values[i].key ? values[i].key : this.cellIndex + i;

                    let valuesItem = this.handlerDataSameMerge(values[i], key);
                    if (valuesItem.fromCellKey && valuesItem.endCellKey) {
                        let mergeRows = 0;

                        if (valuesItem.mergeRows && valuesItem.mergeRows > 1) {
                            mergeRows = valuesItem.mergeRows - 1
                            key = valuesItem.fromCellKey;
                        }
                        else if (valuesItem.mergeRows && valuesItem.mergeRows < -1) {
                            mergeRows = valuesItem.mergeRows + 1
                            key = valuesItem.endCellKey;
                        } else {
                            key = valuesItem.fromCellKey;
                        }

                        if (mergeRows >= 1) {
                            // this.addRowNum += mergeRows;
                        }
                        let beginAddress = this.row.getCell(valuesItem.fromCellKey).address;
                        let endAddress = this.WorkSheet.getRow(this.row._number + mergeRows)
                            .getCell(valuesItem.endCellKey).address;

                        if(style.border){
                            this.row.getCell(valuesItem.fromCellKey).border=style.border
                        }
                        
                        this.WorkSheet.mergeCells(beginAddress, endAddress);
                    } else {
                        key = valuesItem.key ? valuesItem.key : this.cellIndex + i;
                    }

                    cell = this.row.getCell(key);
                    style = await this.handlerMinHeight(style, valuesItem.value, cell._column.width);
                    style['alignment']?style.alignment['wrapText']=true:style['alignment']={wrapText:true}
                    cell.style = style;
                    cell.value = valuesItem.value;
                }

            } else {
                cell = this.row.getCell(this.cellIndex + i);
                styleObj = await this.handlerMinHeight(styleObj, values[i], cell._column.width);
                styleObj['alignment']?styleObj.alignment['wrapText']=true:styleObj['alignment']={wrapText:true}
                cell.style = styleObj;
                cell.value = values[i];
            }
        }
        this.addRowNum++;
    }
    handlerDataSameMerge(valuesItem, key) {
        if (!this.dataSameMergeParams[key]) {
            this.dataSameMergeParams[key] = {
                lastValue: "",
                nextValue: "",
                mergeRows: 1
            }
        }
        if (this.columnCustomPro.dataSameMerge[key] && this.rowType === "bodyRow") {

            if (this.dataRowIndex !== this.sheetData.length - 1) {
                let nextRowData = this.sheetData[this.dataRowIndex + 1];
                let nextValue = ""
                for (let cellItem of nextRowData.values) {
                    if (cellItem.key === key) {
                        nextValue = cellItem.value
                    }
                }
                this.dataSameMergeParams[key].nextValue = nextValue;
            } else {
                this.dataSameMergeParams[key].nextValue = "";
            }

            if (valuesItem.value === this.dataSameMergeParams[key].nextValue) {
                this.dataSameMergeParams[key].mergeRows++;
                this.dataSameMergeParams[key].lastValue = valuesItem.value;
            } else {
                if (this.dataSameMergeParams[key].lastValue === valuesItem.value 
                        && valuesItem.value !== this.dataSameMergeParams[key].nextValue) {
                    valuesItem["fromCellKey"]=key;
                    valuesItem["endCellKey"]=key;
                    valuesItem["mergeRows"]=-this.dataSameMergeParams[key].mergeRows;
                }
            }
        }
        
        return valuesItem;
    }
    handlerMinHeight(style, value, columnWidth) {

        if (style.minHeight && columnWidth) {
            let fontSize = (style.font ? (style.font.size ? style.font.size : 14) : 14);
            let valueLen = (value + '').length * Math.ceil(fontSize * (10 / 85));

            let valueHeight = 2 * (valueLen / columnWidth) * (Math.ceil(fontSize * (10 / 85))) + 3 * Math.ceil(valueLen / columnWidth) + 4
            let height = style.minHeight > valueHeight ? style.minHeight : valueHeight
            delete style.minHeight;

            //重新新賦值行高
            this.row.height = height;
        }

        return style
    }
    async render() {
        await this.initSheet(this.sheet)
        await this.processor();
        return this.WorkBook;
    }
}


class hsExcelUtil {
    constructor({
        OutFileName = "outFile_" + Date.now() + ".xlsx",
        Creator = "",
        Style = {},
        Sheets = []
    }) {
        this.OutFileName = OutFileName;
        this.Creator = Creator;
        this.Style = Style;
        this.Sheets = Sheets;

        this.WorkBook;
    }
    initExcel() {
        if (this.Sheets.length === 0) {
            throw "Sheets length 為 0"
        }
        this.WorkBook = new Excel.Workbook({
            creator: this.creator,
            created: new Date()
        });
    }

    async processor() {
        for (let sheet of this.Sheets) {
            this.WorkBook = await new handlerSheet(
                this.WorkBook,
                sheet,
                this.Style
            ).render();
        }
    }
    creatExcelDoc(type) {
        switch (type) {
            case "file":
                this.WorkBook.xlsx.write(this.OutFileName);
                return this.OutFileName
            case "buff":
                let buff = this.WorkBook.xlsx.writeBuffer()
                return buff
            default:
                this.WorkBook.xlsx.write(this.OutFileName);
                return this.OutFileName
        }
    }
    async render(type) {
        try {
            console.time("excelUtilRunTime:")
            await this.initExcel();
            await this.processor();
            let data = await this.creatExcelDoc(type);
            console.timeEnd("excelUtilRunTime:")
            return { status: 1, data }
        } catch (err) {
            return { status: 0, data: err };
        }
    }
}

module.exports = hsExcelUtil;

/**
 * sheet properties
 *  名称	          默认值	       描述
    tabColor	    undefined	    标签的颜色
    outlineLevelCol	    0	        工作表列大纲级别
    outlineLevelRow	    0	        工作表行大纲级别
    defaultRowHeight	15	        默认行高
    defaultColWidth	(可选)	        默认列宽
    dyDescent	        55	        TBD
 */

/**
 * sheet pageSetup 所有打印的设置
 */

/**
 * sheet views
 * 如1：创建一个隐藏网格线的工作表
 * var sheet = workbook.addWorksheet('My Sheet', {views: [{showGridLines: false}]});
 * 如2：创建一个第一行和列冻结的工作表
 * var sheet = workbook.addWorksheet('My Sheet', {views:[{xSplit: 1, ySplit:1}]});
 */

/**
 * sheet State
 * visible 使工作表可見
 * hidden 使工作表隱藏
 * veryHidden 使工作表隱藏在“隱藏/取消隱藏”對話框中
 */