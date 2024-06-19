/***write by jacky(Chen Hai Quan) 2019/12/20 begin*/
/**write by jacky(Chen Hai Quan) 2019/12/27 end*/

/**update by jacky(Chen Hai Quan) 2019/01/02
 * 1.修復數組表行遍歷字體加粗樣式問題
 * 2.簡化數組數據標記字符
 * 3.去除數組數據標記所佔行
 */

/**
 * update by jacky(Chen Hai Quan) 2020/04/01
 * 1.增加刪除一段範圍的文字功能
 * update by jacky(Chen Hai Quan) 2020/05/11
 * 1.增加删除一整行功能
 * update by jacky(Chen Hai Quan) 2020/05/12
 * 1.增加删除一段节点功能
 */

/**
 * update by jacky(Chen Hai Quan) 2020/05/08
 * 1.增加字符換行功能
 */

/**
 * update by jacky(Chen Hai Quan) 2020/06/09
 * 1.%deleteRow%增加可以刪除多行功能。
 */
/**
 * update by jacky(Chen Hai Quan) 2020/06/10
 * 1.增加%deleteTbl%刪除整個Tbl功能
 *    使用在tbl裡增加一個變量key-value(%deleteTbl%)
 */
"use strict";

/**
  * 功能&使用
  * 1.文本內容標記使用{key},key為傳入data對象的key.
  * 2.數組數據跟進模板生成tbl行列表功能
  *   2-1 傳入的數組需為數組對象模式，如[{name:'jacky',sex:'boy'},{name:'lucy',sex:'girl'}]
  *   2-2 數組文檔標記方式，如：
  *         {#Array1}
  *        
  *         tblItemModeColumn1 {name}    tblItemModeColumn2 {sex}
  *         {#Array1}
  *       Array1-關聯數組數據的data對象的key
  * 3.刪除一段範圍內的文字
  *   標記任意兩個key,value值為"%deleteText%begin"和"%deleteText%end"
  *   將要刪除的文字段包裹在此兩個key標記中
  *   如需要數據判斷來是否刪除文字，那只需要這個兩個key傳''空字符即可
  *   如：
  *      {任意key1}這是要刪除的內容{任意key2}
  * 
  *      data-{key1:'%deleteText%begin',key2:'%deleteText%end'}  
 /**
  * 注意事項
  * 1.避免問題,tag標記請盡可能不要留空格
  */

const fs = require("fs");
const pizzip = require("pizzip");
const path = require("path");

const xmlDom = require("xmldom"),
    DOMParser = xmlDom.DOMParser,
    XMLSerializer = xmlDom.XMLSerializer;
let deleteTextBeginIndex = -1;
let deleteRowBeginIndex = -1;
let deleteRowBeginTcIndex = -1;
let deleteNodeBeginIndex = -1;
class hsDocxUtil {
    constructor(options) {
        this.tagFix = {
            begin: "{",
            end: "}",
            tbl: "#",
        };
        this.xmlNode = {
            t: "w:t",
            tr: "w:tr",
            tbl: "w:tbl",
            p: "w:p",
        };

        this.zip = null;
        this.inputDocName = options["inputFileName"]
            ? options["inputFileName"]
            : "";
        this.outDocName = options["outFileName"] ? options["outFileName"] : "";
        this.data = options["data"] ? options["data"] : {};
        this.contentXmlName = "word/document.xml";

        this.errMsg = "";
        this.status = true;

        /**
         * aPStartNum 數組數據標記開始的行 p標籤位置
         * aPEndNum 數組數據標記結束的行 p標籤位置
         * tblNum 文檔中第幾張表
         * isArrayModeTag 是否為表中的需要遍歷的行模板的tag標記
         */
        this.tagTDom = [];
        this.isGetXmlMode = false;
        this.isArrayModeTag = false;
        this.aPStartNum = -1;
        this.aPEndNum = -1;
        this.tblNum = -1;
        this.tblDomData = [];
        this.tagArrayKey = "";
    }

    setZip() {
        if (this.inputDocName.trim === "") {
            throw "模板文件不能為空";
        }
        let fstat = fs.statSync(path.resolve(__dirname, this.inputDocName));
        if (!fstat.isFile()) {
            throw "模板文件路徑有誤,請檢查";
        }
        this.zip = new pizzip(
            fs.readFileSync(path.resolve(__dirname, this.inputDocName), "binary")
        );
    }

    getXmlFileDom(fileName) {
        if (this.zip["files"][fileName]) {
            return this.zip["files"][fileName].asText();
        } else {
            throw "no files content !";
        }
    }

    stringToXml(str) {
        if (str.charCodeAt(0) === 65279) {
            // 空格
            str = str.substr(1);
        }

        var parser = new DOMParser();
        return parser.parseFromString(str, "text/xml");
    }

    xmlToString(xmlNode) {
        var a = new XMLSerializer();
        return a.serializeToString(xmlNode).replace(/xmlns(:[a-z0-9]+)?="" ?/g, "");
    }

    getPNodeValueObj() {
        let startTagIndex = -1,
            endTagIndex = -1;
        let PNodeValueObj = {};
        for (let j = 0, tlen = this.tagTDom.length; j < tlen; j++) {
            let nodeValue = this.tagTDom[j].childNodes[0].nodeValue;

            if (nodeValue.trim() === "") {
                continue;
            }

            if (nodeValue.indexOf(this.tagFix.begin) !== -1) {
                PNodeValueObj[j] = nodeValue;
                startTagIndex = j;
            }
            if (nodeValue.indexOf(this.tagFix.end) !== -1) {
                PNodeValueObj[j] = nodeValue;
                endTagIndex = j;
            }
            if (
                startTagIndex !== -1 &&
                endTagIndex !== -1 &&
                endTagIndex - startTagIndex > 0
            ) {
                let tagKeyNum = endTagIndex - startTagIndex,
                    startTagNum = startTagIndex;
                for (let n = tagKeyNum; n >= 1; n--) {
                    startTagNum += 1;
                    PNodeValueObj[startTagNum] = this.tagTDom[
                        startTagNum
                    ].childNodes[0].nodeValue;
                }
                startTagIndex = -1;
                endTagIndex = -1;
            }
        }
        return PNodeValueObj;
    }

    getPTagArray(PNodeValueObj) {
        let PTagArray = [],
            PTagArrNum = -1,
            objKey_num = -1;
        for (let objKey in PNodeValueObj) {
            this.tagTDom[objKey].childNodes[0].nodeValue = "";
            this.tagTDom[objKey].childNodes[0].data = "";

            if (PTagArrNum === -1) {
                PTagArray.push({
                    num: objKey,
                    value: PNodeValueObj[objKey].trim(),
                });
                PTagArrNum = 0;
            } else if (PTagArray[PTagArrNum].value.indexOf(this.tagFix.end) === -1) {
                if (
                    PNodeValueObj[PTagArray[PTagArrNum]["num"]].trim().length <=
                    PNodeValueObj[objKey].trim().length
                ) {
                    PTagArray[PTagArrNum]["num"] = objKey;
                }
                PTagArray[PTagArrNum]["value"] += PNodeValueObj[objKey].trim();
            } else if (PTagArray[PTagArrNum].value.indexOf(this.tagFix.end) !== -1) {
                //处理多个标记同一个value========begin
                let obj = PTagArray[PTagArrNum].value.split("").reduce(function (x, y) {
                    return x[y]++ || (x[y] = 1), x;
                }, {});
                if (obj[this.tagFix.end] !== 1) {
                    let PTagArrayItems = PTagArray[PTagArrNum].value.split(
                        this.tagFix.end
                    );
                    if (
                        PTagArrayItems[PTagArrayItems.length - 1].indexOf(
                            this.tagFix.end
                        ) === -1
                    ) {
                        if (
                            PTagArrayItems[PTagArrayItems.length - 1].indexOf(
                                this.tagFix.begin
                            ) === -1
                        ) {
                            PTagArrayItems[PTagArrayItems.length - 2] +=
                                PTagArrayItems[PTagArrayItems.length - 1];
                            PTagArrayItems.splice(PTagArrayItems.length - 1, 1);
                        }
                    }
                    let lastItemObjKey = PTagArray[PTagArray.length - 1].num;
                    PTagArray.splice(PTagArray.length - 1, 1);
                    PTagArrayItems.forEach((item) => {
                        PTagArray.push({
                            num: lastItemObjKey,
                            value: item.trim(),
                        });
                    });
                    PTagArrNum += PTagArrayItems.length - 1;
                }
                //处理多个标记同一个value========end

                if (PTagArray[PTagArrNum].value.indexOf(this.tagFix.end) !== -1) {
                    PTagArray.push({
                        num: objKey,
                        value: PNodeValueObj[objKey].trim(),
                    });
                    PTagArrNum += 1;
                }
            }
        }
        // console.log('PTagArray:',PTagArray)
        return PTagArray;
    }
    getNowValue(val, tagKeyValue) {
        let nowValue = "",
            data = this.data;
        if (this.aPEndNum === -1 && this.aPStartNum === -1) {
            //數組標記內的內容這裡不修改
            nowValue = data.hasOwnProperty(tagKeyValue)
                ? val.replace(
                    this.tagFix.begin + tagKeyValue + this.tagFix.end,
                    data[tagKeyValue]
                )
                : val;
        } else if (this.isGetXmlMode || this.aPEndNum !== -1) {
            nowValue = val;
        } else {
            nowValue = val;
        }
        return nowValue;
    }
    setTextTag(tagTDomIndex, val) {
        let valAry = val.split("<br>");
        for (let i = 0; i < valAry.length; i++) {
            if (i === 0) {
                this.tagTDom[tagTDomIndex].childNodes[0].nodeValue += valAry[i];
                this.tagTDom[tagTDomIndex].childNodes[0].data += valAry[i];
            }
            if (i > 0) {
                this.tagTDom[tagTDomIndex].appendChild(this.stringToXml("<w:br/>"));
                this.tagTDom[tagTDomIndex].appendChild(
                    this.stringToXml("<w:t>" + valAry[i] + "</w:t>")
                );
            }
        }
    }
    deleteText(tagTDomIndex, val) {
        if (val.indexOf("begin") !== -1) {
            deleteTextBeginIndex = tagTDomIndex;
        }
        if (val.indexOf("end") !== -1 && deleteTextBeginIndex !== -1) {
            let endNum = tagTDomIndex;
            for (
                let deleteNum = deleteTextBeginIndex;
                deleteNum <= Number(endNum);
                deleteNum++
            ) {
                this.tagTDom[deleteNum].childNodes[0].nodeValue = "";
                this.tagTDom[deleteNum].childNodes[0].data = "";
            }
            deleteTextBeginIndex = -1;
        }
    }
    deleteRow(contentXml, PTagArray, i, tagTDomIndex, val) {
        if (val.indexOf("begin") !== -1) {
            deleteRowBeginIndex = i;
            deleteRowBeginTcIndex = this.getTcIndex(PTagArray[i]);
        }
        if (val.indexOf("end") !== -1 && deleteRowBeginIndex !== -1) {
            let endNum = i, deleteRowTcIndex = -1;
            for (let i = deleteRowBeginIndex; i <= endNum; i++) {
                deleteRowTcIndex = this.getTcIndex(PTagArray[i]);
                if (
                    ((deleteRowTcIndex !== -1) && (deleteRowTcIndex == deleteRowBeginTcIndex))
                    ||
                    (deleteRowTcIndex === -1)
                ) {
                    let xlmStr = '', tAry = [], tXmlAry = contentXml.getElementsByTagName('w:p')[i].getElementsByTagName('w:t');
                    for (let tIndex = 0; tIndex < tXmlAry.length; tIndex++) {
                        tAry.push('<w:r><w:t> </w:t></w:r>')
                    }
                    xlmStr = '<w:p>' + tAry.join('') + '</w:p>';
                    contentXml.replaceChild(this.stringToXml(xlmStr), PTagArray[i])
                }
            }
            deleteRowBeginIndex = -1;
        }
        return contentXml
    }
    getTcIndex(PTag) {
        let ThisTc = PTag.parentNode;
        let ThisTcNum = -1;
        if (ThisTc.nodeName === 'w:tc') {
            let TcPNodeList = ThisTc.parentNode.getElementsByTagName('w:tc');
            for (let i = 0; i < TcPNodeList.length; i++) {
                if (TcPNodeList[i] == ThisTc) {
                    ThisTcNum = i
                }
            }
        }
        return ThisTcNum;

    }
    deleteTbl(contentXml, PTagArray) {
        let xmlTbl = PTagArray.parentNode.parentNode.parentNode;
        if (xmlTbl.nodeName === "w:tbl") {
            // let tblAllTLen=xmlTbl.getElementsByTagName("w:t").length;
            // let tblStr="",tAry=[];
            // for(let i =0;i<tblAllTLen;i++){
            //     tAry.push("<w:t></w:t>")
            // }
            // tblStr=`<w:tbl><w:tr><w:tc><w:p><w:r>`+tAry.join('')+`</w:r></w:p></w:tc></w:tr></w:tbl>`; 
            // contentXml.replaceChild(this.stringToXml(tblStr),xmlTbl);
            contentXml.removeChild(xmlTbl)
        }
        return contentXml
    }
    deleteNode(tagTDomIndex, val) {
        if (val.indexOf("begin") !== -1) {
            deleteNodeBeginIndex = tagTDomIndex;
        }
        if (val.indexOf("end") !== -1 && deleteNodeBeginIndex !== -1) {
            let endNum = tagTDomIndex;
            let beginPNode = this.tagTDom[deleteNodeBeginIndex].parentNode;
            let endPNode = this.tagTDom[endNum].parentNode;
            let pNode = this.tagTDom[endNum].parentNode.parentNode;
            let pChildNodes = this.tagTDom[endNum].parentNode.parentNode.childNodes;
            let beginPIndex = -1,
                endPIndex = -1;
            for (let i = 0; i < pChildNodes.length; i++) {
                if (pChildNodes.item(i) === beginPNode) {
                    beginPIndex = i;
                }
                if (pChildNodes.item(i) === endPNode) {
                    endPIndex = i;
                }
            }
            for (let j = beginPIndex; j <= endPIndex; j++) {
                pNode.removeChild(pChildNodes.item(beginPIndex));
            }
        }
    }
    renderTextAfetXml(xml) {
        let contentXml = xml;
        let xmlParray = contentXml.getElementsByTagName(this.xmlNode.p);
        if (xmlParray.length === 0) {
            throw "content is empty !";
        }

        for (let i = 0, len = xmlParray.length; i < len; i++) {
            this.tagTDom = xmlParray[i].getElementsByTagName(this.xmlNode.t);
            if (this.tagTDom.length === 0) {
                continue;
            }

            let PNodeValueObj = this.getPNodeValueObj();

            let PTagArray = this.getPTagArray(PNodeValueObj);
            for (let g = 0, glen = PTagArray.length; g < glen; g++) {
                let tagKeyValue = PTagArray[g].value
                    .slice(
                        PTagArray[g].value.indexOf(this.tagFix.begin) + 1,
                        PTagArray[g].value.indexOf(this.tagFix.end)
                    )
                    .trim();
                /**
                 * 1-得到數組數據的開始結束狀態
                 * 2-賦值不是數組數據的其他標記文本
                 * 3-得到數組標記的表序列,改數據表的需要遍歷生成的行模板以及相關數組數據對應的key(對應得到數組數據)結尾重置數組狀態
                 */

                this.setTblBeginEndStatus(tagKeyValue, i);

                let nowValue = this.getNowValue(PTagArray[g].value, tagKeyValue);
                this.setTextTag(PTagArray[g].num, nowValue);

                if (nowValue.indexOf("%deleteText%") !== -1) {
                    this.deleteText(PTagArray[g].num, nowValue);
                }
                if (nowValue.indexOf("%deleteRow%") !== -1) {
                    contentXml = this.deleteRow(contentXml, xmlParray, i, PTagArray[g].num, nowValue);
                }
                if (nowValue.indexOf("%deleteNode%") !== -1) {
                    this.deleteNode(PTagArray[g].num, nowValue);
                }
                if (nowValue.indexOf("%deleteTbl%") !== -1) {
                    contentXml = this.deleteTbl(contentXml, xmlParray[i]);
                }
                this.setTblDomData(tagKeyValue, PTagArray[g].num, contentXml);
            }
        }
        return contentXml;
    }
    setTblBeginEndStatus(tagKeyValue, i) {
        if (tagKeyValue.indexOf(this.tagFix.tbl) !== -1 && this.aPStartNum === -1) {
            this.aPStartNum = i;
            this.isGetXmlMode = true;
        } else if (
            tagKeyValue.indexOf(this.tagFix.tbl) !== -1 &&
            this.aPStartNum !== -1
        ) {
            this.aPEndNum = i;
        }
    }
    setTblDomData(tagKeyValue, PTagArrayGnum, contentXml) {
        if (this.isGetXmlMode) {
            this.tblNum++;
            this.tblDomData.push({
                tblIndex: -1,
                rmode: null,
                rmodeIndex: -1,
                tblArrayKey: tagKeyValue.replace(this.tagFix.tbl, ""),
                preRemoveNode: this.tagTDom[PTagArrayGnum].parentNode.parentNode
                    .parentNode.parentNode,
                nextRemoveNode: null,
                preRemoveNodeIndex: -1,
                nextRemoveNodeIndex: -1,
            });
            this.isGetXmlMode = false;
        }

        if (this.aPEndNum !== -1) {
            let nextRemoveNode = this.tagTDom[PTagArrayGnum].parentNode.parentNode
                .parentNode.parentNode;
            let tblIndex = [].indexOf.call(
                contentXml.getElementsByTagName(this.xmlNode.tbl),
                nextRemoveNode.parentNode
            );
            let rNode = nextRemoveNode.previousSibling;
            let allTrNodes = contentXml.getElementsByTagName(this.xmlNode.tbl)[tblIndex].getElementsByTagName(this.xmlNode.tr);
            let rmodeIndex = [].indexOf.call(allTrNodes, rNode);
            let preRemoveNodeIndex = [].indexOf.call(
                allTrNodes,
                this.tblDomData[this.tblNum].preRemoveNode
            );
            let nextRemoveNodeIndex = [].indexOf.call(allTrNodes, nextRemoveNode);

            this.tblDomData[this.tblNum]["rmode"] = rNode;
            this.tblDomData[this.tblNum]["rmodeIndex"] = rmodeIndex;
            this.tblDomData[this.tblNum]["tblIndex"] = tblIndex;
            this.tblDomData[this.tblNum]["nextRemoveNode"] = nextRemoveNode;
            this.tblDomData[this.tblNum]["preRemoveNodeIndex"] = preRemoveNodeIndex;
            this.tblDomData[this.tblNum]["nextRemoveNodeIndex"] = nextRemoveNodeIndex;

            this.aPStartNum = -1;
            this.aPEndNum = -1;
            this.tagArrayKey = "";
        }
    }

    renderArrayAfterXml(xml) {
        let nowContentXML = xml,
            data = this.data,
            tblDomData = this.tblDomData;
        try {
            for (
                let tblIndex = 0, tblLen = tblDomData.length;
                tblIndex < tblLen;
                tblIndex++
            ) {
                let dataArry = data[tblDomData[tblIndex].tblArrayKey]
                    ? data[tblDomData[tblIndex].tblArrayKey]
                    : [];

                let nextRemoveNodeIsEndR = false;
                let nextRemoveNodeIndex = -1;
                if (
                    dataArry.length !== 0 ||
                    (data[tblDomData[tblIndex].tblArrayKey] &&
                        data[tblDomData[tblIndex].tblArrayKey].length === 0)
                ) {
                    let tblRAarry = nowContentXML
                        .getElementsByTagName(this.xmlNode.tbl)
                    [tblDomData[tblIndex].tblIndex].getElementsByTagName(
                        this.xmlNode.tr
                    );
                    let rmodeIndex = tblDomData[tblIndex].rmodeIndex;
                    let preRemoveNodeIndex = tblDomData[tblIndex].preRemoveNodeIndex;
                    nextRemoveNodeIndex = tblDomData[tblIndex].nextRemoveNodeIndex;

                    if (
                        tblIndex !== 0 &&
                        tblDomData[tblIndex].tblIndex === tblDomData[tblIndex - 1].tblIndex
                    ) {
                        let range = 0;
                        for (let rang = tblIndex - 1; rang >= 0; rang--) {
                            range = range + (data[tblDomData[rang].tblArrayKey].length - 3);
                        }
                        rmodeIndex += range;
                        preRemoveNodeIndex += range;
                        nextRemoveNodeIndex += range;
                    }
                    nowContentXML.removeChild(tblRAarry[rmodeIndex]);
                    nowContentXML.removeChild(tblRAarry[preRemoveNodeIndex]);
                    nowContentXML.removeChild(tblRAarry[nextRemoveNodeIndex]);

                    if (Number(nextRemoveNodeIndex) === tblRAarry.length - 1) {
                        nextRemoveNodeIsEndR = true;
                    }
                }
                for (
                    let dataIndex = 0, dataLen = dataArry.length;
                    dataIndex < dataLen;
                    dataIndex++
                ) {
                    let rmodeStr = this.xmlToString(tblDomData[tblIndex].rmode).replace(
                        ' xmlns:w="' + tblDomData[tblIndex].rmode.namespaceURI + '"',
                        ""
                    );
                    for (let dataItem in dataArry[dataIndex]) {
                        if (
                            rmodeStr.indexOf(
                                this.tagFix.begin + dataItem + this.tagFix.end
                            ) !== -1
                        ) {
                            if (dataArry[dataIndex][dataItem].indexOf("<br>") !== -1) {
                                let strAry = dataArry[dataIndex][dataItem].split("<br>");
                                let newStr = "";
                                for (let strItem = 0; strItem < strAry.length; strItem++) {
                                    if (strItem === strAry.length - 1) {
                                        newStr += "<w:t>" + strAry[strItem] + "<w:t/>";
                                    } else {
                                        newStr += "<w:t>" + strAry[strItem] + "<w:t/><w:br/>";
                                    }
                                }
                                dataArry[dataIndex][dataItem] = newStr;
                            }
                            rmodeStr = rmodeStr.replace(
                                this.tagFix.begin + dataItem + this.tagFix.end,
                                dataArry[dataIndex][dataItem]
                            );
                        }
                    }
                    let oldTblXml = nowContentXML.getElementsByTagName(this.xmlNode.tbl)[
                        tblDomData[tblIndex].tblIndex
                    ];
                    let newTblXml = oldTblXml.cloneNode(true);
                    if (nextRemoveNodeIsEndR) {
                        newTblXml.appendChild(this.stringToXml(rmodeStr));
                    } else {
                        let laterRIndex = nextRemoveNodeIndex - 2 + dataIndex;
                        newTblXml.insertBefore(
                            this.stringToXml(rmodeStr),
                            newTblXml.getElementsByTagName(this.xmlNode.tr)[laterRIndex]
                        );
                    }
                    nowContentXML.replaceChild(newTblXml, oldTblXml);
                }
                nextRemoveNodeIsEndR = false;
            }
        } catch (err) {
            console.log("數組標記有誤:", err);
            this.status = false;
            this.errMsg = "數據標記有誤,請檢查!";
        }

        return nowContentXML;
    }

    render() {
        this.setZip();
        let contentDom = this.getXmlFileDom(this.contentXmlName);
        let contentXml = this.stringToXml(contentDom);
        /**
         * 1-處理非數組標記賦值,同時找出相關數組tbl數據
         * 2-通過得到tbl數據處理數據標記,得到最終doc內容數據
         */

        contentXml = this.renderTextAfetXml(contentXml);
        contentXml = this.renderArrayAfterXml(contentXml);

        return this.xmlToString(contentXml);
    }

    getBuf() {
        let buf = null;
        try {
            const contentXmlstr = this.render();
            if (this.status) {
                this.zip.remove(this.contentXmlName);
                this.zip.file(this.contentXmlName, contentXmlstr, {
                    createFolders: true,
                });

                buf = this.zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
            }
        } catch (error) {
            console.log("error:", error);
            this.errMsg = error;
            this.status = false;
        }

        if (this.status) {
            return {
                status: 1,
                data: buf,
            };
        } else {
            return {
                status: 0,
                data: this.errMsg,
            };
        }
    }
}

module.exports = hsDocxUtil;
