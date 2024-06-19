

/**
 * write by chen hai quan 
 * 微信: Ricardo_CHQ
 */
/**
 * 2023/04/18 chen hai quan,ricardo
 * 1.修復偶爾出現的元素null 報錯
 * 2.修復deleteRow 內容出現不標準符號導致空行 && 移除空標籤
 */
/**
 * 2023/04/18 chen hai quan,ricardo
 * 1.增加footer 頁尾替換字符
 */
/**
 * 2023/06/06 chen hai quan,ricardo
 * 1.修复多个标记同一个value 处理问题，如('{1}{2}' 为一个nodeValue 情况处理)
 */
/**
 * 2023/07/04 chen hai quan,ricardo
 * 1.修复 isMoreTagFix 增加了多余 { ,导致出现 { 问题
 * 2.修复 一整段标记 多个字符 为同一个value,导致出现漏替换出现 { 问题
 */
/**
 * 2023/07/05 chen hai quan,ricardo
 * 1.保留去除 '{' 或 '}' 前后 空格
 */

const fs = require("fs"),
    pizzip = require("pizzip"),
    path = require("path");

const xmlDom = require("xmldom"),
    DOMParser = xmlDom.DOMParser,
    XMLSerializer = xmlDom.XMLSerializer;

class staticMethod {
    constructor() {
        this.tagFix = {
            begin: "{",
            end: "}",
            tbl: "#"
        }
        this.xmlNode = {
            t: "w:t",
            tr: "w:tr",
            tbl: "w:tbl",
            p: "w:p"
        };
        this.contentXmlName = "word/document.xml";
    }

    getXmlFileDom(zip, fileName) {
        if (zip["files"][fileName]) {
            // fs.writeFileSync('./assets/docFile.xml',zip["files"][fileName].asText())
            return zip["files"][fileName].asText();
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

    findParentNode(node, pNodeName, num) {
        let nodeDom = node;
        let i = 0;
        let isFlg = true;
        while (isFlg) {
            if (nodeDom.nodeName === "w:body" || i > num) {
                isFlg = false
                return false
            }
            nodeDom = nodeDom.parentNode;
            if (nodeDom.nodeName === pNodeName) {
                return nodeDom
            }
            i++;
        }
    }
}

class instanceMethod {

    constructor() {
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

        this.pArrayNum = -1;
        this.pArrayData = [];

        this.deleteTextBeginIndex = -1;
        this.deleteRowBeginIndex = -1;
        this.deleteRowBeginTcIndex = -1;
        this.deleteNodeBeginIndex = -1;
        this.deleteContentBeginIndex = -1;
    }

    getPNodeValueObj() {
        let startTagIndex = -1,
            endTagIndex = -1;
        let PNodeValueObj = {};
        for (let j = 0, tLen = this.tagTDom.length; j < tLen; j++) {
            let nodeValue = this.tagTDom[j].childNodes[0].nodeValue;

            if (nodeValue.trim() === "") {
                continue;
            }
            if (nodeValue.indexOf(this.tagFix.begin) !== -1 && startTagIndex === -1) {
                PNodeValueObj[j] = nodeValue;
                startTagIndex = j;
            }
            if (nodeValue.indexOf(this.tagFix.end) !== -1) {
                PNodeValueObj[j] = nodeValue;
                endTagIndex = j;
                if (nodeValue.indexOf(this.tagFix.begin) > nodeValue.indexOf(this.tagFix.end)) {
                    let tagKeyNum = endTagIndex - startTagIndex,
                        startTagNum = startTagIndex;
                    for (let n = tagKeyNum; n >= 1; n--) {
                        startTagNum += 1;
                        PNodeValueObj[startTagNum] = this.tagTDom[startTagNum].childNodes[0].nodeValue;
                    }

                    startTagIndex = j;
                    endTagIndex = -1;
                }
            }
            if (startTagIndex !== -1 && endTagIndex !== -1 && endTagIndex - startTagIndex > 0) {
                let tagKeyNum = endTagIndex - startTagIndex,
                    startTagNum = startTagIndex;
                for (let n = tagKeyNum; n >= 1; n--) {
                    startTagNum += 1;
                    PNodeValueObj[startTagNum] = this.tagTDom[startTagNum].childNodes[0].nodeValue;
                }
                startTagIndex = -1;
                endTagIndex = -1;
            }
        }
        return PNodeValueObj;
    }

    getPTagArray(PNodeValueObj) {
        let PTagArray = [], PTagArrNum = -1, isMoreTagFix = false;
        for (let objKey in PNodeValueObj) {
            this.tagTDom[objKey].childNodes[0].nodeValue = this.tagTDom[objKey].childNodes[0].nodeValue.indexOf(' ') == -1 ? "" : " ";
            this.tagTDom[objKey].childNodes[0].data = this.tagTDom[objKey].childNodes[0].data.indexOf(' ') == -1 ? "" : " ";

            if (isMoreTagFix && PNodeValueObj[objKey].trim() !== '}') {
                PNodeValueObj[objKey] = '{' + PNodeValueObj[objKey]
                isMoreTagFix = false
            }
            if (PTagArrNum === -1) {
                PTagArray.push({
                    num: objKey,
                    value: PNodeValueObj[objKey]
                });
                PTagArrNum = 0;

                if (Object.keys(PNodeValueObj).length == 1) {
                    //处理多个标记同一个value========begin
                    //obj 為每個PTagArray[PTagArrNum].value字符出現的次數 key-字符，value-次數 
                    let obj = PTagArray[PTagArrNum].value.split("").reduce(function (x, y) {
                        return x[y]++ || (x[y] = 1), x;
                    }, {});

                    if (obj[this.tagFix.end] !== 1) {
                        let PTagArrayItems = PTagArray[PTagArrNum].value.split(this.tagFix.end);
                        if (PTagArrayItems[PTagArrayItems.length - 1].indexOf(this.tagFix.end) === -1) {
                            if (PTagArrayItems[PTagArrayItems.length - 1].indexOf(this.tagFix.begin) === -1) {
                                PTagArrayItems[PTagArrayItems.length - 2] += PTagArrayItems[PTagArrayItems.length - 1];
                                PTagArrayItems.splice(PTagArrayItems.length - 1, 1);
                            }
                        }
                        let lastItemObjKey = PTagArray[PTagArray.length - 1].num;
                        PTagArray.splice(PTagArray.length - 1, 1);
                        PTagArrayItems.forEach((item) => {
                            PTagArray.push({
                                num: lastItemObjKey,
                                value: item.indexOf(this.tagFix.end) === -1 ? (item + this.tagFix.end) : item,
                            });
                        });
                        PTagArrNum += PTagArrayItems.length - 1;
                    }
                    //处理多个标记同一个value========end
                }

            } else if (PTagArray[PTagArrNum].value.indexOf(this.tagFix.end) === -1) {
                if (PNodeValueObj[PTagArray[PTagArrNum]["num"]].trim().length <= PNodeValueObj[objKey].trim().length) {
                    PTagArray[PTagArrNum]["num"] = objKey;
                }
                if (
                    (PNodeValueObj[objKey].indexOf(this.tagFix.begin) !== -1 && PNodeValueObj[objKey].indexOf(this.tagFix.end) !== -1)
                    && (PNodeValueObj[objKey].lastIndexOf(this.tagFix.begin) > PNodeValueObj[objKey].lastIndexOf(this.tagFix.end))
                ) {

                    PNodeValueObj[objKey] = PNodeValueObj[objKey].replace(this.tagFix.begin, '')
                    if (!isMoreTagFix) isMoreTagFix = true;
                }
                PTagArray[PTagArrNum]["value"] += PNodeValueObj[objKey];
            } else if (PTagArray[PTagArrNum].value.indexOf(this.tagFix.end) !== -1) {
                //处理多个标记同一个value========begin
                //obj 為每個PTagArray[PTagArrNum].value字符出現的次數 key-字符，value-次數 
                let obj = PTagArray[PTagArrNum].value.split("").reduce(function (x, y) {
                    return x[y]++ || (x[y] = 1), x;
                }, {});

                if (obj[this.tagFix.end] !== 1) {
                    let PTagArrayItems = PTagArray[PTagArrNum].value.split(this.tagFix.end);
                    if (PTagArrayItems[PTagArrayItems.length - 1].indexOf(this.tagFix.end) === -1) {
                        if (PTagArrayItems[PTagArrayItems.length - 1].indexOf(this.tagFix.begin) === -1) {
                            PTagArrayItems[PTagArrayItems.length - 2] += PTagArrayItems[PTagArrayItems.length - 1];
                            PTagArrayItems.splice(PTagArrayItems.length - 1, 1);
                        }
                    }
                    let lastItemObjKey = PTagArray[PTagArray.length - 1].num;
                    PTagArray.splice(PTagArray.length - 1, 1);
                    PTagArrayItems.forEach((item) => {
                        PTagArray.push({
                            num: lastItemObjKey,
                            value: item.indexOf(this.tagFix.end) === -1 ? (item + this.tagFix.end) : item,
                        });
                    });
                    PTagArrNum += PTagArrayItems.length - 1;
                }
                //处理多个标记同一个value========end

                if (PTagArray[PTagArrNum].value.indexOf(this.tagFix.end) !== -1) {
                    PTagArray.push({
                        num: objKey,
                        value: PNodeValueObj[objKey],
                    });
                    PTagArrNum += 1;
                }
            }
        }

        return PTagArray;
    }

    getNowValue(val, tagKeyValue) {
        let nowValue = "",
            data = this.data;
        if (this.aPEndNum === -1 && this.aPStartNum === -1) {
            //數組標記內的內容這裡不修改
            if (tagKeyValue.indexOf('.') === -1) {
                // console.log("tagKeyValue:",tagKeyValue)
                nowValue = data.hasOwnProperty(tagKeyValue)
                    ? val.replace(
                        this.tagFix.begin + tagKeyValue + this.tagFix.end,
                        data[tagKeyValue]
                    )
                    : val;
            } else {
                //為對象賦值
                let objKeys = tagKeyValue.split('.');
                if (data.hasOwnProperty(objKeys[0])) {
                    nowValue = data[objKeys[0]].hasOwnProperty(objKeys[1])
                        ? val.replace(
                            this.tagFix.begin + tagKeyValue + this.tagFix.end,
                            data[objKeys[0]][objKeys[1]]
                        )
                        : val;
                } else {
                    nowValue = val
                }
            }

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
                this.tagTDom[tagTDomIndex].appendChild(this.stringToXml("<w:t>" + valAry[i] + "</w:t>"));
            }
        }
    }

    setTblBeginEndStatus(tagKeyValue, i) {
        if (tagKeyValue.indexOf(this.tagFix.tbl) !== -1 && this.aPStartNum === -1) {
            this.aPStartNum = i;
            this.isGetXmlMode = true;
        } else if (tagKeyValue.indexOf(this.tagFix.tbl) !== -1 && this.aPStartNum !== -1) {
            this.aPEndNum = i;
        }
    }

    setTblDomData(xmlPArray, tagKeyValue, PTagArrayGnum, contentXml) {
        if (this.isGetXmlMode) {

            let preRemoveNode = this.findParentNode(this.tagTDom[PTagArrayGnum], 'w:tr', 4)
            let preRemoveTcNode = this.findParentNode(this.tagTDom[PTagArrayGnum], 'w:tc', 6);
            if (!preRemoveNode || !preRemoveTcNode) {
                this.pArrayNum++;
                this.pArrayData.push({
                    pArrayStartNum: this.aPStartNum,
                    pArrayEndNum: -1,
                    pArrayKey: tagKeyValue.replace(this.tagFix.tbl, "").trim(),
                    modeStr: '',
                    preRemoveDomStr: this.xmlToString(xmlPArray[this.aPStartNum]).replace(' xmlns:w="' + xmlPArray[this.aPStartNum].namespaceURI + '"', ""),
                    nextRemoveDomStr: ''
                })
            } else {
                this.tblNum++;
                this.tblDomData.push({
                    tblIndex: -1,
                    rmode: null,
                    rmodeIndex: -1,
                    tblArrayKey: tagKeyValue.replace(this.tagFix.tbl, "").trim(),
                    preRemoveNode: preRemoveNode,
                    nextRemoveNode: null,
                    preRemoveNodeIndex: -1,
                    nextRemoveNodeIndex: -1,
                });
            }
            this.isGetXmlMode = false;
        }

        if (this.aPEndNum !== -1) {
            let nextRemoveNode = this.findParentNode(this.tagTDom[PTagArrayGnum], 'w:tr', 4);
            let nextRemoveTcNode = this.findParentNode(this.tagTDom[PTagArrayGnum], 'w:tc', 6);
            if (!nextRemoveNode || !nextRemoveTcNode) {
                this.pArrayData[this.pArrayNum].pArrayEndNum = this.aPEndNum;

                this.pArrayData[this.pArrayNum].nextRemoveDomStr = this.xmlToString(xmlPArray[this.aPEndNum])
                    .replace(' xmlns:w="' + xmlPArray[this.aPEndNum].namespaceURI + '"', "");

                let pArrayStartNum = this.pArrayData[this.pArrayNum].pArrayStartNum,
                    pArrayEndNum = this.pArrayData[this.pArrayNum].pArrayEndNum;
                for (let i = pArrayStartNum + 1; i < pArrayEndNum; i++) {
                    this.pArrayData[this.pArrayNum].modeStr += this.xmlToString(xmlPArray[i])
                        .replace(' xmlns:w="' + xmlPArray[i].namespaceURI + '"', "");
                }
            } else {
                let tblIndex = [].indexOf.call(contentXml.getElementsByTagName(this.xmlNode.tbl), nextRemoveNode.parentNode);
                let rNode = nextRemoveNode.previousSibling;
                let allTrNodes = contentXml.getElementsByTagName(this.xmlNode.tbl)[tblIndex].getElementsByTagName(this.xmlNode.tr);
                let rmodeIndex = [].indexOf.call(allTrNodes, rNode);
                let preRemoveNodeIndex = [].indexOf.call(allTrNodes, this.tblDomData[this.tblNum].preRemoveNode);
                let nextRemoveNodeIndex = [].indexOf.call(allTrNodes, nextRemoveNode);

                this.tblDomData[this.tblNum]["rmode"] = rNode;
                this.tblDomData[this.tblNum]["rmodeIndex"] = rmodeIndex;
                this.tblDomData[this.tblNum]["tblIndex"] = tblIndex;
                this.tblDomData[this.tblNum]["nextRemoveNode"] = nextRemoveNode;
                this.tblDomData[this.tblNum]["preRemoveNodeIndex"] = preRemoveNodeIndex;
                this.tblDomData[this.tblNum]["nextRemoveNodeIndex"] = nextRemoveNodeIndex;
            }


            this.aPStartNum = -1;
            this.aPEndNum = -1;
        }
    }

    deleteText(tagTDomIndex, val) {
        if (val.indexOf("begin") !== -1) {
            this.deleteTextBeginIndex = tagTDomIndex;
        }
        if (val.indexOf("end") !== -1 && this.deleteTextBeginIndex !== -1) {
            let endNum = tagTDomIndex;
            for (let deleteNum = this.deleteTextBeginIndex; deleteNum <= Number(endNum); deleteNum++) {
                this.tagTDom[deleteNum].childNodes[0].nodeValue = "";
                this.tagTDom[deleteNum].childNodes[0].data = "";
            }
            this.deleteTextBeginIndex = -1;
        }
    }
    deleteRow(contentXml, PTagArray, i, tagTDomIndex, val) {
        if (val.indexOf("begin") !== -1) {
            this.deleteRowBeginIndex = i;
            this.deleteRowBeginTcIndex = this.getTcIndex(PTagArray[i]);
        }
        if (val.indexOf("end") !== -1 && this.deleteRowBeginIndex !== -1) {
            let endNum = i, deleteRowTcIndex = -1;
            for (let i = this.deleteRowBeginIndex; i <= endNum; i++) {
                deleteRowTcIndex = this.getTcIndex(PTagArray[i]);
                if (
                    ((deleteRowTcIndex !== -1) && (deleteRowTcIndex == this.deleteRowBeginTcIndex))
                    ||
                    (deleteRowTcIndex === -1)
                ) {
                    // let xlmStr = '', tAry = [], tXmlAry = contentXml.getElementsByTagName('w:p')[i].getElementsByTagName('w:t');
                    // for (let tIndex = 0; tIndex < tXmlAry.length; tIndex++) {
                    //     tAry.push('<w:r><w:t> </w:t></w:r>')
                    // }
                    // xlmStr = '<w:p>' + tAry.join('') + '</w:p>';
                    // contentXml.replaceChild(this.stringToXml(xlmStr), PTagArray[i])
                    contentXml.getElementsByTagName("w:body")[0].removeChild(PTagArray[i])
                }
            }
            this.deleteRowBeginIndex = -1;
        }
        return contentXml
    }
    deleteTbl(contentXml, PTagArray) {

        let xmlTbl = this.findParentNode(PTagArray, 'w:tbl');
        if (!xmlTbl) {
            throw 'deleteTbl方法裡獲取tbl錯誤'
        }
        contentXml.removeChild(xmlTbl)
        return contentXml
    }
    deleteNode(tagTDomIndex, val) {
        if (val.indexOf("begin") !== -1) {
            this.deleteNodeBeginIndex = tagTDomIndex;
        }
        if (val.indexOf("end") !== -1 && this.deleteNodeBeginIndex !== -1) {
            let endNum = tagTDomIndex;
            let beginPNode = this.tagTDom[this.deleteNodeBeginIndex].parentNode;
            let endPNode = this.tagTDom[endNum].parentNode;
            let pNode = this.tagTDom[endNum].parentNode.parentNode;
            let pChildNodes = this.tagTDom[endNum].parentNode.parentNode.childNodes;
            let beginPIndex = -1, endPIndex = -1;
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
    deleteContent(contentXml, tagTDomIndex, i, xmlPArray, val) {
        if (val.indexOf("begin") !== -1) {
            this.deleteContentBeginPIndex = i
            this.deleteContentBeginIndex = tagTDomIndex;
        }
        if (val.indexOf("end") !== -1 && this.deleteContentBeginIndex !== -1) {

            let deleteContentEndIndex = tagTDomIndex;
            let beginPNode = xmlPArray[this.deleteContentBeginPIndex];
            let endPNode = xmlPArray[i];

            if (this.deleteContentBeginPIndex === i) { //同一行內容
                let allRNode = beginPNode.getElementsByTagName('w:r');
                let beginRNodeIndex = [].indexOf.call(allRNode, this.tagTDom[this.deleteContentBeginIndex].parentNode);
                let endRNodeIndex = [].indexOf.call(allRNode, this.tagTDom[deleteContentEndIndex].parentNode);
                for (let i = beginRNodeIndex; i <= endRNodeIndex; i++) {
                    beginPNode.removeChild(allRNode[i]);
                }
            } else {
                let beginPNodeParent = beginPNode.parentNode;
                let endPNodeParent = endPNode.parentNode;

                if (beginPNodeParent !== endPNodeParent) {
                    throw 'deleteContent標記有誤'
                }

                let beginIndex = [].indexOf.call(beginPNodeParent.childNodes, beginPNode);

                let endIndex = [].indexOf.call(endPNodeParent.childNodes, endPNode);
                let childNodes = beginPNodeParent.childNodes;
                let delIndex = 0;
                for (let j = beginIndex; j <= endIndex; j++) {
                    if (j == beginIndex) {
                        if (!childNodes[j]) continue;
                        beginPNodeParent.removeChild(childNodes[j])
                    } else {
                        delIndex++;
                        if (!childNodes[j - delIndex]) continue;
                        beginPNodeParent.removeChild(childNodes[j - delIndex])
                    }
                }
            }
        }
        return contentXml
    }
}

class docxUtil extends mix(staticMethod, instanceMethod) {

    constructor({
        inputFileName = "",
        outFileName = "",
        data = {}
    }) {
        super()
        this.zip = null;
        this.inputDocName = inputFileName;
        this.outDocName = outFileName;
        this.data = data;

        this.errMsg = "";
        this.status = true;
    }
    setZip() {
        if (this.inputDocName.trim() === "") {
            throw "模板文件不能為空";
        }
        let fstat = fs.statSync(path.resolve(__dirname, this.inputDocName));
        if (!fstat.isFile()) {
            throw "模板文件路徑有誤,請檢查";
        }
        this.zip = new pizzip(fs.readFileSync(path.resolve(__dirname, this.inputDocName), "binary"));
    }

    renderTextAfterXml(xml) {
        let contentXml = xml;
        let xmlPArray = contentXml.getElementsByTagName(this.xmlNode.p);
        if (xmlPArray.length === 0) {
            throw "content is empty !";
        }

        for (let i = 0, len = xmlPArray.length; i < len; i++) {
            this.tagTDom = xmlPArray[i].getElementsByTagName(this.xmlNode.t);
            if (this.tagTDom.length === 0) {
                continue;
            }

            let PNodeValueObj = this.getPNodeValueObj();
            let PTagArray = this.getPTagArray(PNodeValueObj);
            for (let g = 0, glen = PTagArray.length; g < glen; g++) {
                let tagKeyValue = PTagArray[g].value.slice(
                    PTagArray[g].value.indexOf(this.tagFix.begin) + 1,
                    PTagArray[g].value.indexOf(this.tagFix.end)
                ).trim();
                /**
                 * 代碼塊功能
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

                if (nowValue.indexOf("%deleteNode%") !== -1) {
                    this.deleteNode(PTagArray[g].num, nowValue);
                }
                if (nowValue.indexOf("%deleteTbl%") !== -1) {
                    contentXml = this.deleteTbl(contentXml, xmlPArray[i]);
                }
                if (nowValue.indexOf("%deleteContent%") !== -1) {
                    contentXml = this.deleteContent(contentXml, PTagArray[g].num, i, xmlPArray, nowValue)
                }

                this.setTblDomData(xmlPArray, tagKeyValue, PTagArray[g].num, contentXml);

                if (nowValue.indexOf("%deleteRow%") !== -1) {
                    contentXml = this.deleteRow(contentXml, xmlPArray, i, PTagArray[g].num, nowValue);
                }
            }
        }
        return contentXml;
    }

    renderArrayAfterXml(xml) {
        let nowContentXML = xml,
            data = this.data,
            tblDomData = this.tblDomData;

        try {

            for (let tblIndex = 0, tblLen = tblDomData.length; tblIndex < tblLen; tblIndex++) {

                let dataArry = data[tblDomData[tblIndex].tblArrayKey] ? data[tblDomData[tblIndex].tblArrayKey] : [];

                let nextRemoveNodeIsEndR = false;
                let nextRemoveNodeIndex = -1;

                if (dataArry.length !== 0 || (data[tblDomData[tblIndex].tblArrayKey]
                    && data[tblDomData[tblIndex].tblArrayKey].length === 0)) {
                    let tblRArry = nowContentXML.getElementsByTagName(this.xmlNode.tbl)[tblDomData[tblIndex].tblIndex].getElementsByTagName(this.xmlNode.tr);
                    let rmodeIndex = tblDomData[tblIndex].rmodeIndex;
                    let preRemoveNodeIndex = tblDomData[tblIndex].preRemoveNodeIndex;
                    nextRemoveNodeIndex = tblDomData[tblIndex].nextRemoveNodeIndex;

                    if (tblIndex !== 0 && tblDomData[tblIndex].tblIndex === tblDomData[tblIndex - 1].tblIndex) {
                        let range = 0;
                        for (let rang = tblIndex - 1; rang >= 0; rang--) {
                            range = range + (data[tblDomData[rang].tblArrayKey].length - 3);
                        }
                        rmodeIndex += range;
                        preRemoveNodeIndex += range;
                        nextRemoveNodeIndex += range;
                    }

                    if (this.xmlToString(tblRArry[rmodeIndex]) == '??') {
                        let findTblIndex = (t_index) => {
                            let xml_s = this.xmlToString(tblRArry[t_index])
                            if (xml_s == '??' || (xml_s.indexOf(tblDomData[tblIndex].tblArrayKey) === -1)) {
                                t_index -= 1
                                return findTblIndex(t_index)
                            } else {

                                nowContentXML.removeChild(tblRArry[t_index]);
                                nowContentXML.removeChild(tblRArry[t_index - 1]);
                                nowContentXML.removeChild(tblRArry[t_index - 2]);

                                nextRemoveNodeIndex = t_index;
                            }
                        }
                        findTblIndex(nextRemoveNodeIndex)
                    } else {
                        nowContentXML.removeChild(tblRArry[rmodeIndex]);
                        nowContentXML.removeChild(tblRArry[preRemoveNodeIndex]);
                        nowContentXML.removeChild(tblRArry[nextRemoveNodeIndex]);
                    }

                    if (Number(nextRemoveNodeIndex) === tblRArry.length - 1) {
                        nextRemoveNodeIsEndR = true;
                    }
                }
                for (let dataIndex = 0, dataLen = dataArry.length; dataIndex < dataLen; dataIndex++) {
                    let rmodeStr = this.xmlToString(tblDomData[tblIndex].rmode).replace(' xmlns:w="' + tblDomData[tblIndex].rmode.namespaceURI + '"', "")
                    for (let dataItem in dataArry[dataIndex]) {
                        let reduceFn = () => {
                            if (rmodeStr.indexOf(this.tagFix.begin + dataItem + this.tagFix.end) !== -1) {
                                if (dataArry[dataIndex][dataItem] && dataArry[dataIndex][dataItem].toString().indexOf("<br>") !== -1) {
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
                                rmodeStr = rmodeStr.replace(this.tagFix.begin + dataItem + this.tagFix.end, dataArry[dataIndex][dataItem]);

                                if (rmodeStr.indexOf("%deleteContent%") !== -1) {
                                    // console.log("rmodeStr:",this.stringToXml(rmodeStr))
                                    let xml_rmode = this.stringToXml(rmodeStr);
                                    let thatR_domPs = xml_rmode.getElementsByTagName('w:p');
                                    let delP_begin = -1;
                                    for (let deletePIndex = 0; deletePIndex < thatR_domPs.length; deletePIndex++) {
                                        if (this.xmlToString(thatR_domPs[deletePIndex]).indexOf('%deleteContent%begin') !== -1) {
                                            delP_begin = deletePIndex;
                                        }
                                        if (delP_begin !== -1) {
                                            xml_rmode.removeChild(thatR_domPs[deletePIndex])
                                        }
                                        if (this.xmlToString(thatR_domPs[deletePIndex]).indexOf('%deleteContent%end') !== -1) {
                                            delP_begin = -1;
                                            break;
                                        }
                                    }
                                    rmodeStr = this.xmlToString(xml_rmode)
                                }
                                if (rmodeStr.indexOf(this.tagFix.begin + dataItem + this.tagFix.end) !== -1) {
                                    return reduceFn()
                                }
                            }
                        }
                        reduceFn();
                    }
                    let oldTblXml = nowContentXML.getElementsByTagName(this.xmlNode.tbl)[tblDomData[tblIndex].tblIndex];
                    let newTblXml = oldTblXml.cloneNode(true);
                    if (nextRemoveNodeIsEndR) {
                        newTblXml.appendChild(this.stringToXml(rmodeStr));
                    } else {
                        let laterRIndex = nextRemoveNodeIndex - 2 + dataIndex;
                        newTblXml.insertBefore(this.stringToXml(rmodeStr), newTblXml.getElementsByTagName(this.xmlNode.tr)[laterRIndex]);
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

    renderArrayPAfterXml(xml) {
        let nowContentXMLStr = this.xmlToString(xml),
            data = this.data,
            pArrayData = this.pArrayData;
        for (let i = 0; i < pArrayData.length; i++) {
            if (data.hasOwnProperty(pArrayData[i].pArrayKey) && data[pArrayData[i].pArrayKey]) {
                let allStr = "";
                for (let j = 0; j < data[pArrayData[i].pArrayKey].length; j++) {
                    let modeStr = pArrayData[i].modeStr;
                    for (let key in data[pArrayData[i].pArrayKey][j]) {
                        let reduceFn = () => {
                            modeStr = modeStr.replace(this.tagFix.begin + key + this.tagFix.end, data[pArrayData[i].pArrayKey][j][key])
                            if (modeStr.indexOf(this.tagFix.begin + key + this.tagFix.end) !== -1) {
                                return reduceFn()
                            }
                        }
                        reduceFn()
                    }
                    allStr += modeStr
                }
                nowContentXMLStr = nowContentXMLStr.replace(pArrayData[i].nextRemoveDomStr, '')
                    .replace(pArrayData[i].preRemoveDomStr, '')
                    .replace(pArrayData[i].modeStr, allStr)
            }
        }
        return this.stringToXml(nowContentXMLStr)
    }

    render() {
        if (this.contentXmlName === 'word/document.xml') {//初始setZip
            this.setZip();
        }
        let contentDom = this.getXmlFileDom(this.zip, this.contentXmlName);
        let contentXml = this.stringToXml(contentDom);
        /**
         * 1-處理非數組標記賦值,同時找出相關數組tbl數據
         * 2-通過得到tbl數據處理數據標記,得到最終doc內容數據
         */

        contentXml = this.renderTextAfterXml(contentXml);
        if (this.contentXmlName === 'word/document.xml') { //footer沒有tbl
            contentXml = this.renderArrayAfterXml(contentXml);
        }
        contentXml = this.renderArrayPAfterXml(contentXml);

        let _contentXml = this.xmlToString(contentXml)
        _contentXml = _contentXml.replace(/<w:p\/\>/g, '').replace(/<w:p><w:r><w:t> <\/\w:t><\/\w:r><\/\w:p>/g, '')
        return _contentXml;
    }

    getBuf() {
        let buf = null;
        try {
            //主體內容替換 word/document.xml
            let contentXmlStr = this.render();
            if (this.status) {
                this.zip.remove(this.contentXmlName);
                this.zip.file(this.contentXmlName, contentXmlStr, {
                    createFolders: true,
                });

                //頁尾替換
                for (let fileName in this.zip.files) {
                    if (fileName.indexOf('word/footer') !== -1) {
                        this.contentXmlName = fileName
                        contentXmlStr = this.render();
                        this.zip.remove(this.contentXmlName);
                        this.zip.file(this.contentXmlName, contentXmlStr, {
                            createFolders: true,
                        });
                    }
                }

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

module.exports = docxUtil;

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

