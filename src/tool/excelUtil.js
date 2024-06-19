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
        this.rowType = rowConfig.rowType;
        delete rowConfig.rowType;
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

                    this.dataSameMergeParams[key].mergeRows=1
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

class handlerCss{
    constructor(){
        this.classOtherPro={}
        this.excelCss={}
        this.stylePriority=['rowStyle','filterStyle']//樣式優先,序列越後面,優先級越大
    }

    getExcelCss(className,classData){
        let excelCss={
            "alignment":{},
            "font":{},
            "fill":{},
            "border":{}
        }

        for(let cssName in classData){
            switch(cssName){

                //alignment
                case 'text-algin':
                    //left center right fill justify centerContinuous distributed
                    excelCss['alignment']['horizontal']=classData[cssName]
                break;
                case 'vertical-align':
                    //top middle bottom distributed justify
                    excelCss['alignment']['vertical']=classData[cssName]
                break;
                case 'text-indent':
                    excelCss['alignment']['indent']=classData[cssName]
                break;

                //font
                
                case 'font-size':
                    excelCss['font']['size']=classData[cssName]
                break;
                case 'color':
                    excelCss['font']['color']={
                        "argb":classData[cssName].replace('#','')
                    }
                break;
                case 'font-weight':
                    excelCss['font']['bold']=classData[cssName]==='bold'
                break;
                case 'white-space':
                    excelCss['font']['bold']=classData[cssName]==='bold'
                break;
                
                //fill
                case 'background-color':
                    excelCss['fill']={
                        "type": "pattern",
                        "pattern": "solid",
                        "fgColor": {
                            "argb": classData[cssName].replace('#','')
                        }
                    }
                break;

                //border
                case 'border':
                    excelCss['border']={
                        "top":{
                            "style":"thin",
                            "color":{
                                "argb":classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "left":{
                            "style":"thin",
                            "color":{
                                "argb":classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "right":{
                            "style":"thin",
                            "color":{
                                "argb":classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "bottom":{
                            "style":"thin",
                            "color":{
                                "argb":classData[cssName].split(' ')[2].replace('#','')
                            }
                        }
                    }
                break;
                case 'border-top':
                    excelCss['border']={
                        "top":{
                            "style":"thin",
                            "color":{
                                "argb":classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "left":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "right":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "bottom":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        }
                    }
                break;
                case 'border-left':
                    excelCss['border']={
                        "top":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "left":{
                            "style":"thin",
                            "color":{
                                "argb":classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "right":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "bottom":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        }
                    }
                break;
                case 'border-bottom':
                    excelCss['border']={
                        "top":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "left":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "right":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "bottom":{
                            "style":"thin",
                            "color":{
                                "argb":classData[cssName].split(' ')[2].replace('#','')
                            }
                        }
                    }
                break;
                case 'border-right':
                    excelCss['border']={
                        "top":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "left":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "right":{
                            "style":"thin",
                            "color":{
                                "argb":classData[cssName].split(' ')[2].replace('#','')
                            }
                        },
                        "bottom":{
                            "style":"thin",
                            "color":{
                                "argb":'00'+classData[cssName].split(' ')[2].replace('#','')
                            }
                        }
                    }
                break;
                
                //other pro
                default:
                  if(!this.classOtherPro[className]){this.classOtherPro[className]={}};
                  this.classOtherPro[className][cssName]=classData[cssName];
            }
        }

        //delete empty object
        for(let name in excelCss){
            if(Object.keys(excelCss[name]).length === 0){
                delete excelCss[name]
            }
        }
        this.excelCss=excelCss;
        return excelCss
    }
    getNumStyleValue(styleValue,filterData,value){
        if(typeof filterData['colors'][0][styleValue] === 'string'){
            return Number(value)<Number(filterData.extremum)?filterData['colors'][0][styleValue].replace('#','')
                :(Number(value) === Number(filterData.extremum)?filterData['colors'][1][styleValue].replace('#','')
                    :filterData['colors'][2][styleValue].replace('#',''))
        }
        else if(typeof filterData['colors'][0][styleValue] === 'boolean'){
            return Number(value)<Number(filterData.extremum)?filterData['colors'][0][styleValue]
                :(Number(value) === Number(filterData.extremum)?filterData['colors'][1][styleValue]
                    :filterData['colors'][2][styleValue])
        }
        
    }
    getFilterNumStyle(styleCol,filterData,value){
        if(filterData['colors'] && filterData['colors'].length === 3){
            let color,bg,bold;
            if(typeof filterData['colors'][0] === 'string'){
                color=Number(value)<Number(filterData.extremum)?filterData['colors'][0].replace('#','')
                        :(Number(value) === Number(filterData.extremum)?filterData['colors'][1].replace('#','')
                            :filterData['colors'][2].replace('#',''))
            }
            else if(typeof filterData['colors'][0] === 'object'){
                if(filterData['colors'][0].color){
                    color=this.getNumStyleValue('color',filterData,value)
                }
                if(filterData['colors'][0].bg){
                    bg=this.getNumStyleValue('bg',filterData,value)
                }
                if(filterData['colors'][0].bold){
                    bold=this.getNumStyleValue('bold',filterData,value)
                }
            }
            
            if(color){
                styleCol['font']?styleCol['font']['color']={argb:color}:styleCol['font']={color:{argb:color}};
            }
            if(bold){
                styleCol['font']?styleCol['font']['bold']=bold:styleCol['font']={bold:bold};
            }
            if(bg){
                styleCol['fill']?styleCol['fill']['fgColor']={"argb": bg}
                    :styleCol['fill']={
                        "type": "pattern",
                        "pattern": "solid",
                        "fgColor": {
                            "argb": bg
                        }
                    }
            }
        }

        return styleCol
    }
    getFilterNullStyle(styleCol,filterData){
        if(filterData.nullBg){
            styleCol['fill']={
                "type": "pattern",
                "pattern": "solid",
                "fgColor": {
                    "argb": filterData.nullBg.replace('#','')
                }
            }
        }
        return styleCol
    }
    getFilterStringStyle(styleCol,filterData,value){
        if(filterData.colors){
            for(let i=0;i<filterData.colors.length;i++){
                if(value === filterData.colors[i].value){

                    filterData.colors[i]['color']?styleCol['font']={
                        color:{
                            argb:filterData.colors[i]['color'].replace('#','')
                        },
                        bold:filterData.colors[i]['bold'] || false
                    }:null

                    filterData.colors[i]['bg']?styleCol['fill']={
                        "type": "pattern",
                        "pattern": "solid",
                        "fgColor": {
                            "argb": filterData.colors[i]['bg'].replace('#','')
                        }
                    }:null

                }
            }
        }
        return styleCol
    }
    getFilterAllStyle(styleCol,filterData){
        filterData['color']?styleCol['font']={
            color:{
                argb:filterData['color'].replace('#','')
            },
            bold:filterData['bold'] || false
        }:null

        filterData['bg']?styleCol['fill']={
            "type": "pattern",
            "pattern": "solid",
            "fgColor": {
                "argb": filterData['bg'].replace('#','')
            }
        }:null

        return styleCol
    }
    async getFilterStyle(styleCol,value,key){

        if(!this.columnsFilter[key]){return styleCol}
        let filterData=this.columnsFilter[key];
        if(!filterData.type){return styleCol;}

        switch(filterData.type){
            case 'number':
                if(!isNaN(Number(filterData.extremum))){
                    if(value !== null &&!isNaN(Number(value))){
                        styleCol= await this.getFilterNumStyle(styleCol,filterData,value)
                    }
                    else if(value === null){
                        styleCol = await this.getFilterNullStyle(styleCol,filterData)
                    }
                }
            break;
            case 'string':
                if(value !== null && value !== undefined){
                    styleCol = await this.getFilterStringStyle(styleCol,filterData,value)
                }
                else if(value === null){
                    styleCol = await this.getFilterNullStyle(styleCol,filterData)
                }
            break;
            case 'all':
                if(value !== null && value !== undefined){
                    styleCol = await this.getFilterAllStyle(styleCol,filterData)
                }
                else if(value === null){
                    styleCol = await this.getFilterNullStyle(styleCol,filterData)
                }
            break;
        }
        return styleCol
    }
    getRowStyle(styleCol,value,key){

        return styleCol
    }
    handlerStylePriority(key){
        let priority={ //數字越小，優先級越高。
            'rowStyle':2,
            'filterStyle':1
        }
        switch (true){
            case this.columnsFilter[key] && this.columnsFilter[key].priority:
                priority.filterStyle=this.columnsFilter[key].priority
            case this.rowStyle && this.rowStyle.priority:
                priority.rowStyle=this.rowStyle.priority
        }
        this.stylePriority.sort((a,b)=>{
            return priority[a]-priority[b]
        })
    }
    async getStyleCol(styleCol,value,key){

        await this.handlerStylePriority(key)

        for(let i =0;i<this.stylePriority.length;i++){
            switch (this.stylePriority[i]){
                case 'rowStyle':
                    styleCol= await this.getRowStyle(styleCol,value,key);
                break;
                case 'filterStyle':
                    styleCol= await this.getFilterStyle(styleCol,value,key)
                break;
            }
        }

        return styleCol
    }
}

class handlerColumnsAndHeaderRowsData{
    constructor(){
        this.titleRows=[]
        this.columnRows=[]
        this.columnsWidth={}
        this.dataSameMerge={}
        this.reduceRes={}
        this.lastReduceRes={}
        this.diffAry={}
        this.columns=[]
        this.columnsFilter={}
        this.columnStyle={}
    }

    async getHeight(rowItem){
        if(rowItem.style && rowItem.style.height){
            return rowItem.style.height
        }
        else if(rowItem.styleClass){
            let height=20;//默認行高20
            for(let i =0;i<rowItem.styleClass.length;i++){
                await this.getExcelCss(rowItem.styleClass[i])
                if(this.classOtherPro[rowItem.styleClass[i]] && this.classOtherPro[rowItem.styleClass[i]].height){
                    height=this.classOtherPro[rowItem.styleClass[i]].height;
                    break;
                }
            }
            return height
        }
        else{
            return 20
        }
    }

    async getTitleRows(){
        let titleRows=[],excelHeader=this.excelHeader;
        for(let i =0;i<excelHeader.length;i++){
            let height=await this.getHeight(excelHeader[i])
            let val=excelHeader[i].caption;
            if(typeof val === 'string'){
                val=val.replace('<br/>',`
`)
            }
            titleRows.push({
                rowConfig:{height:height},
                styleClass:excelHeader[i].styleClass || [],
                values:[
                    {
                        value:val,
                        fromCellKey:this.columns[0].key,
                        endCellKey:this.columns[this.columns.length-1].key
                    }
                ]
            })
        }
        this.titleRows=titleRows
    }

    async getCellKeys(columns,rowIndex){
        let columnAry=columns,CellKeys=[],reduceNum=1;
        let reduceFn= async (columnAry)=>{
            for(let i =0;i<columnAry.length;i++){
                if(columnAry[i].dataField){
                    CellKeys.push(columnAry[i].dataField)
                }
                if(columnAry[i].column && columnAry[i].column.length !== 0){
                    reduceNum++;
                   await reduceFn(columnAry[i].column)
                }
            }
        }
        await reduceFn(columnAry)
        return {
            cellKeys:CellKeys,
            reduceNum:reduceNum,
            rowIndex:rowIndex
        }
    }
    setLastRowDifField(reduceRes,rowIndex){
        if(reduceRes.cellKeys.length !== this.lastReduceRes.cellKeys.length){
            let lastCellKeys=new Set(this.lastReduceRes.cellKeys);
            let cellKeys=new Set(reduceRes.cellKeys);
            let diff=[...new Set([...lastCellKeys].filter(v=>!cellKeys.has(v)))]; //差值
            if(!this.diffAry[rowIndex]){
                this.diffAry[rowIndex]=[]
            }
            this.diffAry[rowIndex]=diff
        }
    }
    getColumnRowsKeys(columnRowValues){
        let keys=[]

        for(let i = 0;i<columnRowValues.length;i++){
            if(columnRowValues[i].key){
                keys.push(columnRowValues[i].key)
            }else{
                let beginKeyIndex=this.columns.indexOf(columnRowValues[i].fromCellKey);
                let endKeyIndex=this.columns.indexOf(columnRowValues[i].endCellKey);
                keys=keys.concat(this.columns.slice(beginKeyIndex,endKeyIndex+1))
            }
        }
        return keys
    }

    async rowsReduceFn(excelColumn,rowIndex){
        if(!this.columnRows[rowIndex]){
            let height=await this.getHeight(excelColumn[0])
            this.columnRows.push({
                rowConfig:{height:height},
                styleClass:excelColumn[0].styleClass || [],
                values:[]
            })
        }

        for(let i=0;i<excelColumn.length;i++){

            let val=excelColumn[i].caption;
            if(typeof val === 'string'){
                val=val.replace('<br/>',`
`)
            }
            if(!excelColumn[i].dataField){
                
                if(excelColumn[i].column && excelColumn[i].column.length !== 0){
                    let reduceRes= await this.getCellKeys(excelColumn[i].column,rowIndex);
                    
                    let cellKeys=reduceRes.cellKeys;
                    this.columns=this.columns.concat(cellKeys)
                    this.reduceRes=reduceRes;

                    if(cellKeys.length === 1){
                        this.columnsWidth[cellKeys[0]]=excelColumn[i].column[0].width || 10;
                        this.dataSameMerge[cellKeys[0]]=excelColumn[i].column[0].dataSameMerge || false
                        this.columnRows[rowIndex].values.push({
                            value:val,
                            key:cellKeys[0]
                        })
                    }else{
                        if(cellKeys[0] !== cellKeys[cellKeys.length-1]){
                            this.columnRows[rowIndex].values.push({
                                value:val,
                                fromCellKey:cellKeys[0],
                                endCellKey:cellKeys[cellKeys.length-1]
                            })
                        }
                    }
                    
                    let index=rowIndex+1;
                  await this.rowsReduceFn(excelColumn[i].column,index)
                }
            }else{
                this.columnsWidth[excelColumn[i].dataField]=excelColumn[i].width || 10;
                this.dataSameMerge[excelColumn[i].dataField]=excelColumn[i].dataSameMerge || false
                this.columns.push(excelColumn[i].dataField)
                if(excelColumn[i].filter){
                    this.columnsFilter[excelColumn[i].dataField]=excelColumn[i].filter
                }
                if(excelColumn[i].columnStyle){
                    this.columnStyle[excelColumn[i].dataField]=excelColumn[i].columnStyle
                }

                this.columnRows[rowIndex].values.push({
                    key:excelColumn[i].dataField,
                    value:val
                })
            }
        }
    }
    async getColumnsAndHeaderRows(){
        let excelColumn=this.excelColumn;
        if(excelColumn.length == 0){return this.columnRows}

        await this.rowsReduceFn(excelColumn,0)

        this.columns=[...new Set(this.columns)];

        await this.setColumnRowsMerge()

        await this.setColumnsWith()
    }
    setColumnRowsMerge(){
        let columnRows=this.columnRows.reverse();
        let mergeNum=1;
        for(let i=0;i<columnRows.length;i++){
            let keys=this.getColumnRowsKeys(columnRows[i].values);
            if(keys.length === this.columns.length){
                break;
            }
            keys=new Set(keys)
            let lastRowKeys=this.getColumnRowsKeys(columnRows[i+1].values);
            let diff=lastRowKeys.filter(item=>!keys.has(item));
            if(diff.length !== 0){
                mergeNum++;
                for(let j=0;j<diff.length;j++){
                    for(let k=0;k<columnRows[i+1].values.length;k++){
                        if(diff[j] === columnRows[i+1].values[k].key){
                            columnRows[i+1].values[k]['fromCellKey']=diff[j]
                            columnRows[i+1].values[k]['endCellKey']=diff[j]
                            columnRows[i+1].values[k]['mergeRows']=mergeNum
                            delete columnRows[i+1].values[k].key
                        }
                    }
                }
            }
        }
        this.columnRows=this.columnRows.reverse();

    }
    setColumnsWith(){
        let columns=this.columns;
        for(let i =0;i<columns.length;i++){
            columns[i]={
                key:columns[i],
                width:this.columnsWidth[columns[i]] || 10,
                dataSameMerge:this.dataSameMerge[columns[i]] || false
            }
        }
        this.columns=columns;
    }

}

class handlerBodyRowsData{
    constructor(){
        this.bodyRows=[]
        this.groupValues={}
        this.groupData={}
    }

    handlerGroup(){
        let excelGroup=this.excelGroup;
        let reduceFn=(excelGroup)=>{
            if(!excelGroup.groupField){return false}
            this.groupData[excelGroup.groupField]={};
            if(excelGroup.summary){
                for(let i = 0;i<excelGroup.summary.length;i++){

                    this.groupData[excelGroup.groupField][excelGroup.summary[i].dbField]={
                        sumType:excelGroup.summary[i].sumType,
                        displayFormat:excelGroup.summary[i].displayFormat || '',
                        value:excelGroup.summary[i].value || ''
                    }

                }
            }
            if(excelGroup.group){
                excelGroup=excelGroup.group
                return reduceFn(excelGroup)
            }
        }
        reduceFn(excelGroup)
    }
    async getGroupRowRes(excelItemData){

        let keys=Object.keys(this.groupData),addRows=[];
        for(let i =0;i<keys.length;i++){
            if(!this.groupValues[keys[i]] || (this.groupValues[keys[i]].value !== excelItemData[keys[i]]) 
                || (addRows.length !== 0) ){
                addRows.push({
                    sumRowData:this.groupData[keys[i]],
                    outlineLevel:i+1,
                    value:excelItemData[keys[i]]
                })
                if(!this.groupValues[keys[i]]){
                    this.groupValues[keys[i]]={}
                }
                this.groupValues[keys[i]]['value']=excelItemData[keys[i]]
                this.groupValues[keys[i]]['outlineLevel']=i+1;
                this.groupValues[keys[i]]['sumRowData']=this.groupData[keys[i]];
            }
        }
        return {addRows}
    }

    getRowNum(local_level,groupMarks,key,sumType){
        let _sumType=sumType.toUpperCase()
        let num=0,total=0;
        let start=groupMarks[local_level].startRows;
        let end=groupMarks[local_level].endRows;
        
        for(let j=start;j<end;j++){
            if(!isNaN(Number(this.excelData[j][key]))){
                total+=Number(this.excelData[j][key])
                switch (_sumType){
                    case "SUM":
                       num+=Number(this.excelData[j][key])
                    break;
                    case "MAX":
                       num = Number(this.excelData[j][key])>num?Number(this.excelData[j][key]):num
                    break;
                    case "SIN":
                        num = Number(this.excelData[j][key])<num?Number(this.excelData[j][key]):num
                    break;
                }
            }
        }
        switch (_sumType){
            case 'AVERAGE':
                num=total/(i-start)
            break;
            case 'COUNT':
                num=i-start;
            break;
        }

        return num;
    }
    
    async getBodyRows(){
        let bodyRows=[],excelData=this.excelData;
        let bodyRowNum=-1,groupRowRes={};
        let groupMarks={}
        await this.handlerGroup()

        for(let i=0;i<excelData.length;i++){
            
            groupRowRes = await this.getGroupRowRes(excelData[i]);

            //set group-summary ====begin
            if(i !== 0){
                for(let h=groupRowRes.addRows.length-1;h>=0;h--){

                    groupMarks[groupRowRes.addRows[h].outlineLevel]['endRows']=i;
                    bodyRows.push({
                        rowConfig:{
                            outlineLevel:groupRowRes.addRows[h].outlineLevel
                        },
                        styleClass:['fontBold'],
                        values:['']
                    })
                    bodyRowNum++;

                    for(let key in groupRowRes.addRows[h].sumRowData){
                        let num=0;
                        if(groupRowRes.addRows[h].sumRowData[key].sumType === 'text'){
                            num = groupRowRes.addRows[h].sumRowData[key].value || ''
                        }else{
                            num=await this.getRowNum(groupRowRes.addRows[h].outlineLevel,groupMarks,key,groupRowRes.addRows[h].sumRowData[key].sumType)
                        }
                        bodyRows[bodyRowNum].values.push({
                            key:key,
                            value:num,
                            styleCol:{
                                numFmt:groupRowRes.addRows[h].sumRowData[key].displayFormat || ''
                            }
                        })
                    }
                }
                
            }
            //set group-summary ====end

            //set groupRow =======begin
            for(let k =0;k<groupRowRes.addRows.length;k++){
                let outlineLevel=groupRowRes.addRows[k].outlineLevel
                groupMarks[outlineLevel]={
                    startRows:i,
                    endRows:0
                }
                bodyRows.push({
                    rowConfig:{},
                    styleClass:['_groupClass'],
                    styleCol:{
                        alignment:{
                            indent:outlineLevel>1?outlineLevel+1:outlineLevel
                        }
                    },
                    values:[{
                        value:groupRowRes.addRows[k].value,
                        fromCellKey:this.columns[0].key,
                        endCellKey:this.columns[this.columns.length-1].key
                    }]
                })
                bodyRowNum++;
                if(groupRowRes.addRows[k].outlineLevel !== 1){
                    bodyRows[bodyRowNum].rowConfig['outlineLevel']=groupRowRes.addRows[k].outlineLevel-1
                }
            }
            //set groupRow =======end

            //set bodyRow && other data handler ====begin
            bodyRows.push({
                rowConfig:{
                    rowType:'bodyRow'
                },
                styleClass:[],
                values:[]
            })
            bodyRowNum++;
            let groupKeys=Object.keys(this.groupValues);
            if(groupKeys.length !== 0){
                bodyRows[bodyRowNum].rowConfig['outlineLevel']=this.groupValues[groupKeys[groupKeys.length-1]]['outlineLevel'];
            }
            for(let j=0;j<this.columns.length;j++){
                let styleCol={};
                    styleCol = await this.getStyleCol(styleCol,excelData[i][this.columns[j].key],this.columns[j].key,(j===this.columns.length-1))
                let val=excelData[i][this.columns[j].key];
                if(typeof val === 'string'){
                    val=val.replace('<br/>',`
`);
                }
                bodyRows[bodyRowNum].values.push({
                    key:this.columns[j].key,
                    value:val,
                    styleCol:styleCol
                })
            }
            //set bodyRow && other data handler ====end

            //set last group-summary ===begin
            if(i === (excelData.length-1)){
                for(let z=groupKeys.length-1;z>=0;z--){
                    groupMarks[this.groupValues[groupKeys[z]].outlineLevel]['endRows']=i+1;
                    bodyRows.push({
                        rowConfig:{
                            outlineLevel:this.groupValues[groupKeys[z]].outlineLevel
                        },
                        styleClass:['fontBold'],
                        values:[]
                    })
                    bodyRowNum++;

                    for(let key in this.groupValues[groupKeys[z]].sumRowData){
                        let num=0;
                        if(this.groupValues[groupKeys[z]].sumRowData[key].sumType === 'text'){
                            num=this.groupValues[groupKeys[z]].sumRowData[key].value;
                        }else{
                            num=await this.getRowNum(this.groupValues[groupKeys[z]].outlineLevel,groupMarks,key,this.groupValues[groupKeys[z]].sumRowData[key].sumType)
                        }
                        bodyRows[bodyRowNum].values.push({
                            key:key,
                            value:num,
                            styleCol:{
                                numFmt:this.groupValues[groupKeys[z]].sumRowData[key].displayFormat || ''
                            }
                        })
                    }
                }
            }
            //set last group-summary ===end

        }
        this.bodyRows=bodyRows;
    }
}

class handlerFooter{
    constructor(){ 
        this.footerRows=[]
    }
    getFooterSummaryNum(summaryItem){
        let sumType=summaryItem.sumType.toUpperCase(),
            key=summaryItem.dbField;
        let num=0,total=0;
        if(sumType === 'TEXT'){
            num=summaryItem.value || '';
            return num;
        }
        if(sumType === 'COUNT'){
            num=this.excelData.length;
            return num;
        }
        
        for(let i=0;i<this.excelData.length;i++){
            let _num=this.excelData[i][key];
            if(!isNaN(Number(_num))){
                total+=Number(_num)
                switch(sumType){
                    case 'SUM':
                        num+=Number(_num)
                    break;
                    case 'MAX':
                        num=num<Number(_num)?Number(_num):num
                    break;
                    case 'SIN':
                        num=num>Number(_num)?Number(_num):num
                    break;
                }
            }
        }
        if(sumType === 'AVERAGE'){
            return total/this.excelData.length
        }
        return num
    }
    async getFooterRows(){
        let footerRows=this.footerRows,
            excelFooter=this.excelFooter;
        for(let i =0;i<excelFooter.length;i++){
            footerRows.push({
                rowConfig:{},
                styleClass:['fontBold'],
                values:[]
            })
            if(excelFooter[i].summary){
                for(let j=0;j<excelFooter[i].summary.length;j++){
                    let num=this.getFooterSummaryNum(excelFooter[i].summary[j])
                    footerRows[i].values.push({
                        key:excelFooter[i].summary[j].dbField,
                        value:num,
                        styleCol:{
                            numFmt:excelFooter[i].summary[j].displayFormat || ''
                        }
                    })
                }
            }
            else if(excelFooter[i].value){   
                let height=await this.getHeight(excelFooter[i])
                footerRows[i].rowConfig['height']=height
                footerRows[i].values.push(excelFooter[i])
            }
        }
    }
}

class ExcelUtil extends mix(
    handlerCss,
    handlerColumnsAndHeaderRowsData,
    handlerBodyRowsData,
    handlerFooter
) {
    constructor(
        {
            export:exportObj={
                location:"",
                filename:""
            },
            header:excelHeader=[],
            column:excelColumn=[],
            group:excelGroup={},
            data:excelData=[],
            footer:excelFooter=[],
            style:excelClass={},
            rowStyle=false
        }
    ){
        super()
        this.OutFileName=exportObj.location+exportObj.filename;
        this.excelColumn=excelColumn;
        this.excelHeader=excelHeader;
        this.excelData=excelData;
        this.excelGroup=excelGroup;
        this.excelFooter=excelFooter;
        this.excelClass=excelClass;
        this.rowStyle=rowStyle;
        
        this.Sheets=[{
            SheetName:exportObj.filename+"_"+Date.now(),
            Columns:[],
            startRowIndex:1,
            startCellIndex:1,
            sheetData:[]
        }]

        //內置Class
        this.builtInClass={
            darkTitleClass:{
                'text-algin':'center',
                'vertical-align':'middle',
                "font-size":16,
                "font-weight":'bold',
                "color":"#ffffff",
                "background-color":'#161616',
                height:50
            },
            grayHeaderClass:{
                border:'1px solid #000000',
                "text-algin":'center',
                "font-weight":'bold',
                'background-color':'#e7e6e6',
                'vertical-align':'middle'
            },
            _groupClass:{
                'background-color':'#D0CECE',
                'border':'1px solid #444444',
                'font-weight':'bold'
            },
            fontBold:{
                'font-weight':'bold'
            },
            defaultMarkBox:{
                'border':'1px solid #444444',
                "vertical-align":'top',
                height:66
            }
        }
    }

    async getSheets(){
        await this.getColumnsAndHeaderRows()
        await this.getTitleRows()
        await this.getBodyRows()
        await this.getFooterRows()

        this.Sheets[0].Columns=this.columns;
        this.Sheets[0].sheetData=this.titleRows.concat(this.columnRows,this.bodyRows,this.footerRows)
        return this.Sheets
    }
    getClass(){
        let excelClass=this.excelClass;
        Object.assign(excelClass,this.builtInClass)

        for(let className in excelClass){
            excelClass[className]=this.getExcelCss(className,excelClass[className])
        }
        
        return excelClass
    }
    
    async render(type){
       let style= await this.getClass()
       let sheets=await this.getSheets()
       return new hsExcelUtil({
            OutFileName:this.OutFileName,
            Style:style,
            Sheets:sheets
        }).render(type)
    }
}

function mix(...mixins) {
    class Mix {
        constructor(...ags) {
            for (let mixin of mixins) {
                copyProperties(this, new mixin(ags)); // 拷贝实例属性
            }
        }
    }
    for (let mixin of mixins) {
        copyProperties(Mix, mixin); // 拷贝静态属性
        copyProperties(Mix.prototype, mixin.prototype); // 拷贝原型属性
    }
    return Mix
}
function copyProperties(target, source) {
    for (let key of Reflect.ownKeys(source)) {
        if (key !== 'constructor'
            && key !== 'prototype'
            && key !== 'name'
        ) {
            let desc = Object.getOwnPropertyDescriptor(source, key);
            Object.defineProperty(target, key, desc);
        }
    }
}

module.exports=ExcelUtil;