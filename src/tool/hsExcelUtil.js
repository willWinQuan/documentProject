/**
 * update by Chen Hai Quan ---2020/04/09
 * 1.增加大綱級別可折疊分類的配置功能。
 */
/**
 * update by Chen Hai Quan ---2020/04/14
 * 1.增加設置最小高度配置項。
 * update by Chen Hai Quan ---2020/04/16
 * 1.最小高度配置項自適應高度問題修復。
 */
/**
 * update by Chen Hai Quan ---2020/04/15
 * 1.修復bodyRow數據為空數組生成報表報錯問題。
 */

const Excel = require('exceljs');

class hsExcelUtil {
    constructor(excelJson) {
        this.excelJson = excelJson;
        this.workbook = {};
        
    }
    /**异步创建Excel,并返回一个buffer，返回类型为一个promise */
    creatExcelBufferAsync = async () => {
        let data,status;
        try {
            let book = this.createWordBook();
            data = await book.xlsx.writeBuffer();
            status=1;
        } catch (error) {
            console.log("createExcelAsync err", error);
             status=0;
             data=error;
            // throw new Error("creatExcelBuffer error");
        }
        return {status:status,data:data};
    }

    createWordBook = () => {
        try {
            this.workbook = new Excel.Workbook();
            if (this.excelJson.creator) {
                this.workbook.creator = this.excelJson.creator;
                this.workbook.lastModifiedBy = this.excelJson.creator;
            }
            this.workbook.created = new Date();
            this.workbook.modified = new Date();
            this.workbook.lastPrinted = new Date();
            this.workbook.properties.date1904 = true;
            this.workbook.calcProperties.fullCalcOnLoad = true;
            this.drawExcelContent();
        } catch (error) {
            console.log("createWordBook err", error);
            throw error;
        }
        return this.workbook;
    }

    drawExcelContent = () => {
        try {
            if (this.excelJson.sheets && this.excelJson.sheets.length > 0) {
                for (let sheet of this.excelJson.sheets) {
                    let sheetC = new ExcelSheet(this.workbook, sheet, this.excelJson.styles);
                    sheetC.createSheet();
                }
            }
        } catch (error) {
            console.log("createWordBook err", error);
            throw error;
        }
    }
}

class ExcelSheet {
    colSerialNumbers = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
    styleList = {};
    worksheet = {};
    sheet = {};
    sheetColsKey = {};
    header = {};
    title = {};
    footer = {};
    constructor(workbook, sheet, styleList) {
        this.workbook = workbook;
        this.sheet = sheet;
        this.worksheet = this.workbook.addWorksheet(this.sheet.name);
        this.styleList = styleList;
        this.bodyEndRow=0;
        this.contentMaxHeight=0;
    }

    createSheet = () => {
        try {
            this.createHeader();
            this.createTitle();
            this.createBody();
            this.createFooter();
        } catch (error) {
            console.log("createSheet err", error);
            throw error;
        }
    }

    createTitle = () => {
        try {
            if (this.sheet && this.sheet.title && this.sheet.title.rowsData) {
                this.title = new Title();
                this.title.setStartRowIndex(this.getTitleStartRowIndex());
                this.title.setEndRowIndex(this.getTitleEndRowIndex());
                for (let rowD of this.sheet.title.rowsData) {
                    let rowIndex = this.title.getStartRowIndex() + (rowD.rowIndex - 1);
                    let row = this.worksheet.getRow(rowIndex);
                    // let fromColNum = this.getColIndexByKey(rowD.fromColKey);
                    // let endColNum = this.getColIndexByKey(rowD.endColKey);
                    let rendCell;
                    let fromCol, endCol, fromColAddr, endColAddr;
                    if (rowD.fromColKey) {
                        let cell = row.getCell(rowD.fromColKey)
                        cell.value = rowD.value;
                        fromCol = cell;
                        fromColAddr = fromCol._address;
                        rendCell = cell;
                    }
                    if (rowD.endColKey) {
                        let cell = row.getCell(rowD.endColKey)
                        cell.value = rowD.value;
                        endCol = cell;
                        endColAddr = endCol._address;
                        rendCell = cell;
                    }
                    if (fromColAddr && endColAddr) {
                        this.worksheet.mergeCells((fromColAddr + ":" + endColAddr));
                        fromCol.value = rowD.value;
                        rendCell = fromCol;
                    }
                    if (rendCell) {
                        this.renderColStyle(rowD, this.styleList, rendCell);
                    }
                    this.commonRenderRowHeight(row, rowD.rowHeight);
                    this.renderColHeight(rowIndex, rowD, this.styleList);
                }
            }
        } catch (error) {
            console.log("createTitle err", error);
            throw error;
        }
    }

    createHeader = () => {
        try {
            if (this.sheet && this.sheet.header && this.sheet.header.cols) {
                this.header = new Header();
                this.header.setStartRowIndex(this.getHeaderStartRowIndex());
                this.header.setEndRowIndex(this.getHeaderEndRowIndex());
                this.header.setStartColIndex(this.sheet.startColIndex);
                this.header.setDeep(this.handleGetTreeDeep(this.sheet.header.cols))
                let preCol = null;
                for (let col of this.sheet.header.cols) {
                    col.rowIndex = this.header.getStartRowIndex();
                    col.colIndex = preCol ? (preCol.colIndex) : (this.header.getStartColIndex() - 1);
                    this.addCol(col);
                    preCol = col;
                    if (col.subCol) {
                        preCol.colIndex += this.getColHasCols(col).count//获取当前列的最底层子节点个数
                    }
                }

                let columns = [];
                let keyObj = {};
                for (let k in this.sheetColsKey) {//遍历所有的子节点的key
                    let a = { key: k };
                    columns.push(a);
                    keyObj[k] = { "keyIndex": this.sheetColsKey[k].keyIndex }
                }
                this.header.setKeys(keyObj);
                this.worksheet.columns = columns;//设置表头的key
                this.renderHeaderWidth(this.styleList);
            }
        } catch (error) {
            console.log("createSheet err", error);
            throw error;
        }
    }

    createBody = () => {
        try {
            let rowIndex = this.getBodyStartRowIndex();
            if (this.sheet && this.sheet.body && this.sheet.body.rowsData) {
                rowIndex --;
                if(this.sheet.body.rowsData.length === 0){
                  this.bodyEndRow=rowIndex;
                }
                for (let row of this.sheet.body.rowsData) {
                    if (row['rowMsg'] && row['rowMsg']['rowType'] == 'group') {//数据行类型为group
                        this.addGroupRowData(this.worksheet, this.sheet, row, rowIndex);//创建group行数据
                        rowIndex++;
                        continue;
                    }
                   
                    let rowCount = this.addBodyRowData(this.worksheet, this.sheet, row, rowIndex);//创建普通数据行
                    // console.log("rowCount",rowIndex,rowCount);
                    rowIndex = rowIndex + rowCount;
                    this.bodyEndRow = rowIndex-1;
                }

            }
        } catch (error) {
            console.log("createBody err", error);
            throw error;
        }
    }

    /**添加group行数据 */
    addGroupRowData = (worksheet, sheet, row, rowIndex) => {
        try {
            let groupCont = {};
            let groupObj = {};
            let firstValIsGroupName = 0;
            for (let key in row['data']) {//取第一个有值的key作为，group显示的值
                groupCont[key] = row['data'][key].value;
                if (row['data'][key].value && firstValIsGroupName == 0) {//取第一个有值的作为group需要显示的值
                    groupObj = row['data'][key];
                    firstValIsGroupName++;
                    break;
                }
            }
            let keys = [];
            for (let k in this.sheetColsKey) {//需要添加空列，否则row.eachCell无法遍历全部的key
                keys.push(k);
                if (!groupCont[k]) {
                    groupCont[k] = null;
                }
            }
            let sum = {};
            sum.fromColKey = keys[0];
            sum.endColKey = keys[keys.length - 1];
            let gRow = worksheet.getRow(rowIndex);
            let rendCell;
            let startCellAddr, endCellAddr, startCell, endCell;
            if (sum.fromColKey) {
                let cell = gRow.getCell(sum.fromColKey)
                cell.value = groupObj.value;
                startCellAddr = cell._address;
                rendCell = cell;
                startCell = cell;
            }
            if (sum.endColKey) {
                let cell = gRow.getCell(sum.endColKey)
                cell.value = groupObj.value;
                endCellAddr = cell._address;
                rendCell = cell;
            }
            if (endCellAddr && startCellAddr) {//合并单元格
                worksheet.mergeCells((startCellAddr + ":" + endCellAddr));
                startCell.value = groupObj.value;
                rendCell = startCell;
            }
            if (rendCell) {
                if (sheet.body.baseStyle) {
                    if (sheet.body.baseStyle[rendCell._column._key]) {
                        this.renderColStyle(sheet.body.baseStyle[rendCell._column._key], this.styleList, rendCell);
                    }
                }
                this.renderColStyle(groupObj, this.styleList, rendCell);
            }
            if (sheet.body.baseStyle && sheet.body.baseStyle['rowHeight']) {
                this.commonRenderRowHeight(gRow, sheet.body.baseStyle['rowHeight']);
            }
            if (row['rowMsg'] && row['rowMsg']['rowHeight']) {
                this.commonRenderRowHeight(gRow, row['rowMsg']['rowHeight']);
            }
            this.renderColHeight(rowIndex, groupObj, this.styleList);
        } catch (error) {
            console.log("addGroupRowData err", error);
            throw error
        }
    }

    /**添加一行普通的数据行 row为数据行的内容json*/
    addBodyRowData = (worksheet, sheet, row, rowIndex) => {
        try {
            let incRow = 1;
            let rowObj = worksheet.getRow(rowIndex + incRow-1);
            for (let key in row['data']) {
                if (row['data'][key].constructor == Array) {
                    let cnt = this.addBodySubRowData(worksheet, sheet, row['data'][key], rowIndex);
                    incRow = incRow > cnt ? incRow : cnt;
                }
                else
                  rowObj.getCell(key).value = row['data'][key].value;
                //    if(key === "pendingItem"){
                //     console.log(rowObj.getCell(key)._address)
                //     let cellAddress=rowObj.getCell(key)._address;
                //      rowObj.getCell(key).value = {

                //          result:'LEFT('+cellAddress+',LEN('+cellAddress+')-11)&CHAR(10)&RIGHT('+cellAddress+',11)'
                //      }
                //     //  rowObj.getCell('pendingRepDate').value={
                //     //      formula:'LEFT('+cellAddress+',LEN('+cellAddress+')-11)&CHAR(10)&RIGHT('+cellAddress+',11)',
                //     //      result:row['data'][key].value
                //     //  }
                //    }
                //    if(row['data'][key].formula){
                //     rowObj.getCell(key).formula=row['data'][key].formula;
                //     rowObj.getCell(key).result=row['data'][key].value;
                //    }else{
                //     rowObj.getCell(key).value = row['data'][key].value;
                //    }
                    
            }
            if (sheet.body.baseStyle && sheet.body.baseStyle['rowHeight']) {
                this.commonRenderRowHeight(rowObj, sheet.body.baseStyle['rowHeight']);
            }
            if (row['rowMsg'] && row['rowMsg']['rowHeight']) {
                this.commonRenderRowHeight(rowObj, row['rowMsg']['rowHeight']);
            }
            if(row['rowMsg'] && row['rowMsg']['outlineLevel']){
                worksheet.getRow(rowIndex + incRow-1).outlineLevel=row['rowMsg']['outlineLevel'];
                worksheet.properties.outlineProperties = {
                    summaryBelow: false,
                    summaryRight: false
                };
            }
            /**render row baseStyle */
            if (sheet.body.baseStyle) {
                rowObj.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    let baseStyle = sheet.body.baseStyle;
                    // console.log("row each cell base style",baseStyle[cell._column._key]);
                    if (baseStyle[cell._column._key]) {
                        this.renderColStyle(baseStyle[cell._column._key], this.styleList, cell);
                        this.renderColHeight(rowIndex, baseStyle[cell._column._key], this.styleList);
                    }
                })
            }
            rowObj.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                // console.log("row each cell style",row[cell._column._key]);
                this.renderColStyle(row['data'][cell._column._key], this.styleList, cell);
                this.renderColHeight(rowIndex, row['data'][cell._column._key], this.styleList);
            })
            if (incRow > 1) {
                rowObj.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    // console.log("row each cell style",row[cell._column._key]);
                    if (row['data'][cell._column._key].constructor != Array) {
                        let MStartCell = cell._address;
                        let MEndCell = worksheet.getRow(rowIndex + incRow-1).getCell(cell._column._key)._address;
                        // console.log(MStartCell + ":" + MEndCell);
                        worksheet.mergeCells(MStartCell + ":" + MEndCell);
                    }
                })
            }
            
            return incRow;

        } catch (error) {
            console.log("addBodyRowData err", error);
            throw error;
        }
    }

    /**添加数组的数据行 row为数据行的内容json*/

    addBodySubRowData = (worksheet, sheet, data, rowIndex) => {
        // console.log("DDD",JSON.stringify(data));
        try {
            let incRow = 0;
            for (let row of data) {
                incRow++;
                let rowObj = worksheet.getRow(rowIndex + incRow-1);
                // console.log("Test",JSON.stringify(row));
                for (let key in row['data']) {
                    rowObj.getCell(key).value = row["data"][key].value;
                    if (row["data"][key].colStyle) {
                        this.commonRenderColStyle(row["data"][key].colStyle, rowObj.getCell(key));
                    }
                }
                if (sheet.body.baseStyle && sheet.body.baseStyle['rowHeight']) {
                    this.commonRenderRowHeight(rowObj, sheet.body.baseStyle['rowHeight']);
                }
                if (row['rowMsg'] && row['rowMsg']['rowHeight']) {
                    this.commonRenderRowHeight(rowObj, row['rowMsg']['rowHeight']);
                }
                if(row['rowMsg'] && row['rowMsg']['outlineLevel']){
                    worksheet.getRow(rowIndex + incRow-1).outlineLevel=row['rowMsg']['outlineLevel'];
                }
                /**render row baseStyle 
                if (sheet.body.baseStyle) {
                    rowObj.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                        let baseStyle = sheet.body.baseStyle;
                        // console.log("row each cell base style",baseStyle[cell._column._key]);
                        if (baseStyle[cell._column._key]) {
                            this.renderColStyle(baseStyle[cell._column._key], this.styleList, cell);
                            this.renderColHeight(rowIndex, baseStyle[cell._column._key], this.styleList);
                        }
                    })
                }*/
                /**
                rowObj.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    // console.log("row each cell style",row[cell._column._key]);
                    this.renderColStyle(row[cell._column._key], this.styleList, cell);
                    this.renderColHeight(rowIndex, row[cell._column._key], this.styleList);
                })*/
            }
            incRow = incRow > 0 ? incRow : 0;
            
            return incRow;
        } catch (error) {
            console.log("addBodySubRowData err", error);
            throw error;
        }
    }

    createFooter = () => {
        try {
            if (this.sheet && this.sheet.footer && this.sheet.footer.rowsData) {
                this.footer = new Footer();
                this.footer.setStartRowIndex(this.getFooterStartRowIndex());
                this.footer.setEndRowIndex(this.getFooterEndRowIndex());
                for (let rowD of this.sheet.footer.rowsData) {
                    let rowIndex = this.footer.getStartRowIndex() + (rowD.rowIndex - 1);
                    let row = this.worksheet.getRow(rowIndex);
                    // let fromColNum = this.getColIndexByKey(rowD.fromColKey);
                    // let endColNum = this.getColIndexByKey(rowD.endColKey);
                    let rendCell;
                    let fromCol, endCol, fromColAddr, endColAddr;
                    if (rowD.fromColKey) {
                        let cell = row.getCell(rowD.fromColKey)
                        cell.value = rowD.value;
                        fromCol = cell;
                        fromColAddr = fromCol._address;
                        rendCell = cell;
                    }
                    if (rowD.endColKey) {
                        let cell = row.getCell(rowD.endColKey)
                        cell.value = rowD.value;
                        endCol = cell;
                        endColAddr = endCol._address;
                        rendCell = cell;
                    }
                    if (fromColAddr && endColAddr) {
                        this.worksheet.mergeCells((fromColAddr + ":" + endColAddr));
                        fromCol.value = rowD.value;
                        rendCell = fromCol;
                    }
                    if (rendCell) {
                        this.renderColStyle(rowD, this.styleList, rendCell);
                    }
                    this.commonRenderRowHeight(row, rowD.rowHeight);
                    this.renderColHeight(rowIndex, rowD, this.styleList);
                }
            }
        } catch (error) {
            console.log("createFooter err", error);
            throw error;
        }
    }

    addCol = (col) => {
        try {
            if (!col.subCol) {//D类型的col
                let ckey = this.getNextColSeriaNumber(col.colIndex, col.rowIndex) + "";//获取cell的地址，如A4
                let cell = this.worksheet.getCell(ckey);
                cell.value = col.colName;
                cell["colKey"]=col.colKey;
                col.address = ckey;
                col.keyIndex = col.colIndex;
                let dept = this.header.getDeep() + (this.header.getStartRowIndex() - 1);//获取表头的深度
                if ((dept) > col.rowIndex) {//如果当前列所在的行数少于表头的深度，则需要向下合并
                    let dkey = this.getNextColSeriaNumber(col.colIndex, dept) + "";
                    this.worksheet.mergeCells((ckey + ":" + dkey));
                }
                col.colIndex++;
                this.sheetColsKey[col.colKey] = col;
                this.renderColStyle(col, this.styleList, cell);
                this.renderColHeight(col.rowIndex, col, this.styleList)
                return col;
            }
            if (col.subCol) {//该列有子列
                let ckey = this.getNextColSeriaNumber(col.colIndex, col.rowIndex);
                let cell = this.worksheet.getCell(ckey);
                cell.value = col.colName;
                cell["colKey"]=col.colKey;
                this.renderColStyle(col, this.styleList, cell);
                this.renderColHeight(col.rowIndex, col, this.styleList)
                let tempColIndex = col.colIndex;
                let ban = col.colIndex;
                for (let c of col.subCol) {
                    c.rowIndex = col.rowIndex + 1;//存在子列，则行号加1
                    c.colIndex = tempColIndex;
                    let t = this.addCol(c);
                    let te = this.getColHasCols(c).count;
                    tempColIndex += te;
                }
                let s = this.getColHasCols(col).count - 1;
                // console.log("ban+s",ban+s,ban,s)
                let tkey = this.getNextColSeriaNumber(ban + s, col.rowIndex);//取横向合并的地址
                this.worksheet.mergeCells((ckey + ":" + tkey));
            }
        } catch (error) {
            console.log("addCol err", error);
            throw error;
        }
    }
    getColIndexByKey = (key) => {
        if (key) {
            let keys = this.header.getKeys();
            if (keys) {
                let col = keys[key];
                if (col) {
                    return col['keyIndex'];
                }
            }
        }
        return null;
    }
    getHasRows = (rowsData) => {
        let num = 0;
        if (rowsData) {
            for (let row of rowsData) {
                if (row.rowIndex > num) {
                    num = row.rowIndex;
                }
            }
        }
        return num;
    }
    getTitleStartRowIndex = () => {
        let index = 1;
        if (this.sheet.startRowIndex) {
            index = this.sheet.startRowIndex;
        }
        return index;
    }
    getTitleEndRowIndex = () => {
        let index = this.getTitleStartRowIndex();
        let titleDeep = 0;
        if (this.sheet.title) {
            titleDeep = this.getHasRows(this.sheet.title.rowsData);
        }
        index = index + (titleDeep - 1);
        return index;
    }
    getHeaderStartRowIndex = () => {
        let index = this.getTitleEndRowIndex() + 1;
        if (this.sheet.header && this.sheet.header.marginTopRows) {
            index += (this.sheet.header.marginTopRows);
        }
        return index;
    }

    getHeaderEndRowIndex = () => {
        let index = this.getHeaderStartRowIndex();
        let headerDeep = (this.sheet.header && this.sheet.header.cols) ? this.handleGetTreeDeep(this.sheet.header.cols) : 0;
        index = index + headerDeep - 1;
        return index;
    }

    getBodyStartRowIndex = () => {
        let index = this.getHeaderEndRowIndex() + 1;
        if (this.sheet.body && this.sheet.body.marginTopRows) {
            index += (this.sheet.body.marginTopRows);
        }
        return index;
    }

    getBodyEndRowIndex = () => {
        let index = this.getBodyStartRowIndex();
        let bodyDeep = (this.sheet.body && this.sheet.body.rowsData) ? this.sheet.body.rowsData.length : 0;
        index = index + bodyDeep - 1;
        return index;
    }

    getFooterStartRowIndex = () => {
        // let index = this.getBodyEndRowIndex() + 1;
        let index = this.bodyEndRow +1;
        if (this.sheet.footer && this.sheet.footer.marginTopRows) {
            index += (this.sheet.footer.marginTopRows);
        }
        return index;
    }

    getFooterEndRowIndex = () => {
        let index = this.getFooterStartRowIndex();
        let footerDeep = 0;
        if (this.sheet.footer) {
            footerDeep = this.getHasRows(this.sheet.footer.rowsData);
        }
        index = index + (footerDeep - 1);
        return index;
    }

    /**渲染表头的宽度，cell的宽度只在头部设置里起效果 */
    renderHeaderWidth = (styles) => {
        if (this.sheetColsKey) {
            for (let key in this.sheetColsKey) {
                let colObj = this.sheetColsKey[key];
                let column = this.worksheet.getColumn(key);
                if (colObj.colStyleClass && styles) {
                    for (let styleClassName of colObj.colStyleClass) {
                        let styleO = styles[styleClassName];
                        if (styleO.width) {
                            column.width = styleO.width;
                        }
                    }
                }
                if (colObj.colStyle) {
                    if(colObj.colStyle.width){
                        column.width = colObj.colStyle.width;
                    }
                }
            }
        }
    }
    renderColHeight = (rowIndex, colS, styles) => {
        let row = this.worksheet.getRow(rowIndex);
        let width=0,minHeight=0;
        if (colS.colStyleClass && styles) {
            for (let styleClassName of colS.colStyleClass) {
                let styleO = styles[styleClassName];
                if (styleO.height) {
                    this.commonRenderRowHeight(row, styleO.height)
                }
                if (styleO.minHeight) {
                    minHeight=styleO.minHeight;
                    if(styleO.width){
                        width=styleO.width;
                        
                        this.setRowMinHeight(row,styleO,colS.colKey);
                    }
                }
            }
        }
        if (colS.colStyle && colS.colStyle.height) {
            this.commonRenderRowHeight(row, colS.colStyle.height);
        }
        if(((colS.colStyle && colS.colStyle.minHeight) || minHeight) && (width || colS.colStyle.width)){
            if(!colS.colStyle.width){
               colS.colStyle["width"]=width;
            }
            // console.log("width:",colS.colStyle.width)
            // for(let item of row._cells){
            //     if(item._address === "AD2"){
            //       console.log("sata:",colS.colStyle["width"])
            //     }
            // }
            
            if(!colS.colStyle.minHeight){
                colS.colStyle["minHeight"]=minHeight;
            }
            
            this.setRowMinHeight(row,colS.colStyle,colS.colKey);
        }
        
    }
    setRowMinHeight =(row ,styleO,colKey)=>{
        //計算內容高
        let fontSize=styleO.font?
        styleO.font.size?styleO.font.size:11
        :11;
        for(let item of row._cells){
            if(item.colKey === colKey){
                let contentHeight=
                Math.ceil((item._value.model.value.length*fontSize*0.75)/styleO.width)*(fontSize/2.5)+5;
                if(this.contentMaxHeight<contentHeight){
                   this.contentMaxHeight=Math.floor(contentHeight);
                } 
                break;
            }
        }
        this.commonRenderRowHeight(row,
            this.contentMaxHeight>styleO.minHeight?this.contentMaxHeight:styleO.minHeight);
        
    }
    commonRenderRowHeight = (row, height) => {
        if (row) {
            row.height = height;
        }
    }
    /**样式渲染 */
    renderColStyle = (col, styles, cell) => {
        if (col.colStyleClass && styles) {
            for (let styleClassName of col.colStyleClass) {
                let style = styles[styleClassName];
                // console.log("style name",style);
                this.commonRenderColStyle(style, cell);
                
            }
        }
        if (col.colStyle) {
            this.commonRenderColStyle(col.colStyle, cell);
        }
    }
    /**通用样式渲染 */
    commonRenderColStyle = (style, cell) => {
        if (style.font) {
            cell.font = style.font;
        }
        if (style.alignment) {
            cell.alignment = style.alignment;
        }
        if (style.border) {
            let border = {};
            border.top = style.border.top ? style.border.top : (cell.style.border ? cell.style.border.top : null)
            border.left = style.border.left ? style.border.left : (cell.style.border ? cell.style.border.left : null)
            border.right = style.border.right ? style.border.right : (cell.style.border ? cell.style.border.right : null)
            border.bottom = style.border.bottom ? style.border.bottom : (cell.style.border ? cell.style.border.bottom : null)
            cell.border = border;
        }
        if (style.fill) {
            cell.fill = style.fill;
        }
        if(style.numFmt){
            cell.numFmt=style.numFmt;
        }
    }

    /**获取表头某一列的底层节点的个数 */
    getColHasCols = (col) => {
        if (!col.subCol) {
            col.count = 1;
            return col;
        }
        if (col.subCol) {
            let cot = 0;
            for (let c of col.subCol) {
                let co = this.getColHasCols(c);
                if (co) {
                    cot += co.count;
                }
            }
            col.count = cot;
            return col;
        }
    }
    /**获取单元格的地址 */
    getNextColSeriaNumber = (numIndex, rowIndex) => {
        let fist = "";
        let sec = "";
        let f = parseInt((numIndex) / 26);
        if (f) {
            fist = this.colSerialNumbers[f - 1];
        }
        let s = numIndex % 26;
        sec = this.colSerialNumbers[s];
        return fist + sec + rowIndex;
    }
    /**获取树节点的深度 */
    handleGetTreeDeep = (fileHeader) => {
        let deep = 0;
        fileHeader.forEach((item) => {
            if (item.subCol) {
                deep = Math.max(deep, this.handleGetTreeDeep(item.subCol) + 1);
            } else {
                deep = Math.max(deep, 1);
            }
        });
        return deep;
    }

}
class Title {
    startRowIndex = 0;
    endRowIndex = 0;
    constructor() {

    }
    setStartRowIndex = (index) => {
        this.startRowIndex = index;
    }

    getStartRowIndex = () => {
        return this.startRowIndex;
    }

    setEndRowIndex = (index) => {
        this.endRowIndex = index;
    }

    getEndRowIndex = () => {
        return this.endRowIndex;
    }
}
class Footer {
    startRowIndex = 0;
    endRowIndex = 0;
    constructor() {

    }
    setStartRowIndex = (index) => {
        this.startRowIndex = index;
    }

    getStartRowIndex = () => {
        return this.startRowIndex;
    }

    setEndRowIndex = (index) => {
        this.endRowIndex = index;
    }

    getEndRowIndex = () => {
        return this.endRowIndex;
    }
}
class Header {
    deep = 0;
    startRowIndex = 0;
    endRowIndex = 0;
    startColIndex = 0;
    endColIndex = 0;
    keys = {};
    constructor() {

    }
    setDeep = (deep) => {
        this.deep = deep;
    }
    getDeep = () => {
        return this.deep;
    }
    setStartRowIndex = (index) => {
        this.startRowIndex = index;
    }

    getStartRowIndex = () => {
        return this.startRowIndex;
    }

    setEndRowIndex = (index) => {
        this.endRowIndex = index;
    }

    getEndRowIndex = () => {
        return this.endRowIndex;
    }

    setStartColIndex = (index) => {
        this.startColIndex = index;
    }

    getStartColIndex = () => {
        return this.startColIndex;
    }

    setEndColIndex = (index) => {
        this.endColIndex = index;
    }

    getEndColIndex = () => {
        return this.endColIndex;
    }

    setKeys = (keys) => {
        this.keys = keys;
    }
    getKeys = () => {
        return this.keys;
    }

}
module.exports = hsExcelUtil;