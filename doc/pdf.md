
# pdf工具類說明文檔

  **依賴pdfkit通過傳入json數據生成pdf文檔。**  

  > link:[pdfkit-npm.md](pdfkit-npm.md  "pdfkit")
  
---
 **log:**<br/>
  > begin-write by Chen Hai Quan (jacky) 2020/04/09  
  > end-write by Chen Hai Quan (jacky) 2020/04/17  
---

# 大綱
  <ul>
    <li>
      <a href="#basic">1-基本用法</a>
      <ul>
        <li>
          <a href="#basicResiter">1.1-文檔註冊配置</a>
        </li>
        <li>
          <a href="#basicStructure">1.2-文檔數據結構</a>
        </li>
        <li>
          <a href="#basicFeatures">1.3-工具功能概覽</a>
        </li>
      </ul>
    </li>
    <li>
      <a href="#text">2-Text</a>
    </li>
    <li>
      <a href="#image">3-Image</a>
    </li>
    <li>
      <a href="#checkbox">4-Checkbox</a>
    </li>
    <li>
      <a href="#grid">5-Grid</a>
    </li>
    <li>
      <a href="#styleTalk">6-樣式詳解</a>
      <ul>
         <li>
           <a href="#textStyle">6.1-Text樣式</a>
         </li>
         <li>
           <a href="#imageStyle">6.2-Image樣式</a>
         </li>
         <li>
           <a href="#checkboxStyle">6.3-Checkbox樣式</a>
         </li>
         <li>
           <a href="#gridStyle">6.4-Grid樣式</a>
         </li>
      </ul>
    </li>
    <li>
      <a href="#practicalTips">7-實用技巧</a>
    </li>
  </ul>

## <a id="basic">1-基本用法</a>
  
  ```javascript
    let doc = new hsPDFUtil({
       outFileName:'D:/Henderson/hsRWCS/UploadFiles/documention/outFile_'+Date.now()+".pdf",
       "Space": {
            "top": 16,
            "left": 46,
            "bottom": 30,
            "right": 30
        },
        "Style": {
            "listLable": {
                "isContinued": true
            },
        },
        "Header": [],
        "Body": [
          {
            "Type": "Image",
            "X": -16,
            "Y": 0,
            "ColStyle": {
                "isContinued": true
            }
          },
          {
            "Type": "Text",
            "Obj": {
                "Y": 40,
                "X": -80,
                "Value": "建築部-完成工程報告",
                "ColStyle": {
                    "fontSize": 18,
                    "align": "center",
                    "font": "Dengb"
                }
            }
          },
          {
            Type:"Checkbox",
            Obj:{ 
              ColStyle:{
                margins:{
                  left:6
                }
              },
              CheckData:[
                {
                  isChecked:true,
                  value:'欠缺保養'    
                }
              ]
            }
          },
          {
            Type:"Grid",
            Obj:{   
              Y:6,
              Title:{},
              Columns:{
                data:[
                 {
                  Value:'管工工數',
                  ColStyle:{
                    width:60
                  },
                  {
                    Value:'管工工數',
                    ColStyle:{
                      width:60
                    }   
                  }   
                 }
                ],
                keys:["key1","key2"]
              },
              GridData:[
                {
                   key1:"testChq",
                   key2:"testCHQASF"
                }
              ]
            }
          }
        ],
        "Footer":[]
    });

    //輸出包含文件名稱的文件路徑。
    let res=doc.render();
    let fileName;
    if(res.status){
       filName=res.data.fileName;
    }
  ```

  ### <a id="basicResiter"> &emsp; 1.1-文檔註冊配置</a>

  ```javascript
    new hsPDFUtil({
      outFileName:"可選-輸出文檔路徑&名稱",
      Author:"可選-文檔作者",
      Subject:"可選-主題",
      Keywords:"可選-要素/關鍵字",
      Space:{
        left:"可選-文檔內容左邊距",
        right:"可選-文檔內容右邊距",
        top:"可選-文檔內容上邊距",
        bottom:"可選-文檔內容下邊距"
      },
      font:[//可選
        "需要註冊的字體文件名稱" //需把字體文件放入默認靜態字體文件夾裡 -默認字體為Deng
        ...
      ]
    })
  ```
  > 附：pdfkit默認字體，無需再註冊  
    >> `'Courier'`  
    >> `'Courier-Bold'`  
    >> `'Courier-Oblique'`  
    >> `'Courier-BoldOblique'`  
    >> `'Helvetica'`  
    >> `'Helvetica-Bold'`  
    >> `'Helvetica-Oblique'`  
    >> `'Helvetica-BoldOblique'`  
    >> `'Symbol'`  
    >> `'Times-Roman'`  
    >> `'Times-Bold'`  
    >> `'Times-Italic'`  
    >> `'Times-BoldItalic'`  
    >> `'ZapfDingbats'`  
    
  ### <a id="basicStructure">&emsp;1.2-文檔數據結構</a>
    
  **除文檔註冊配置類以外**
  
  ```javascript
    new hsPDFUtil({
       "Style":{}, //樣式class集
       "Header": [],//文檔頁頭部設置
       "Body": [], //文檔主體內容
       "Footer":[]  //文檔頁尾部設置
    });
  ```

  ### <a id="basicFeatures">&emsp;1.3-工具功能概覽</a>
  
  &emsp;**文檔主體內容配置根據數組的順序遍歷渲染數組中的每個類型對象，<br/>&emsp;類型可為在對象中Type字段中設置，目前可設置類型為Text/Image/Checkbox/label(Text的簡化)。<br/>&emsp;每個類型有對應的數據字段需提供項(Text-value字段,Image-url字段....)，<br/>&emsp;每個類型都可以選擇性的設置StyleName(array-填寫style對象有的樣式class名稱)、ColStyle(object-填寫樣式key-value形式的屬性配置)來作為該類型的樣式。<br/>&emsp;每個類型通過設置X,Y字段數值來控制與前一類型的內容的間距。<br/>&emsp;樣式中的padding屬性是一個無視其他類型的內容佔位而可穿透擴大的無敵屬性。**

## <a id="text">2-Text</a>

```javascript
  {
    "Type":"Text",
    "Obj":{
      "Y":"number(可負數)-可選（默認為0）-與前一類型內容Y軸間隔（無前一類型內容與頁邊界間隔）",
      "X":"number(可負數)-可選（默認為0）-與前一類型內容X軸間隔（無前一類型內容與頁邊界間隔",
      "Value":"string-必需項-需渲染的內容",
      "ColStyle":"object-可選-樣式key-value形式屬性配置(優先級高於StyleName設置的class類樣式)",
      "StyleName":"array-可選-填寫style對象有的樣式class名稱"
    }
  }
```

## <a id="image">3-Image</a>

```javascript
  {
    "Type":"Image",
    "Obj":{
      "Y":"number(可負數)-可選（默認為0）-與前一類型內容Y軸間隔（無前一類型內容與頁邊界間隔）",
      "X":"number(可負數)-可選（默認為0）-與前一類型內容X軸間隔（無前一類型內容與頁邊界間隔",
      "url":"string-必需項-圖片路徑/base64",
      "ColStyle":"object-可選-樣式key-value形式屬性配置(優先級高於StyleName設置的class類樣式)",
      "StyleName":"array-可選-填寫style對象有的樣式class名稱"
    }
  }
```

## <a id="checkbox">4-Checkbox</a>

```javascript
  {
    "Type":"Checkbox",
    "Obj":{
      "Y":"number(可負數)-可選（默認為0）-與前一類型內容Y軸間隔（無前一類型內容與頁邊界間隔）",
      "X":"number(可負數)-可選（默認為0）-與前一類型內容X軸間隔（無前一類型內容與頁邊界間隔",
      "ColStyle":"object-可選-樣式key-value形式屬性配置(優先級高於StyleName設置的class類樣式)",
      "StyleName":"array-可選-填寫style對象有的樣式class名稱",
      "CheckData":[ //array-必填-checkbox組的內容項
        {
          "isChecked":"boolean-必填-是否勾選",
          "value":"string-必填-check內容",
          "ColStyle":"object-可選-樣式key-value形式屬性配置(優先級高於StyleName設置的class類樣式)-優於數組外面",
          "StyleName":"array-可選-填寫style對象有的樣式class名稱-優於數組外面"
        }
      ]
    }
  }
```

## <a id="grid">5-Grid</a>

```javascript
  {
    "Type":"Text",
    "Obj":{
      "Y":"number(可負數)-可選（默認為0）-與前一類型內容Y軸間隔（無前一類型內容與頁邊界間隔）",
      "X":"number(可負數)-可選（默認為0）-與前一類型內容X軸間隔（無前一類型內容與頁邊界間隔",
      "Title": { //表格標題 可選
        "value":"string-標題",
        "ColStyle":"object-可選-樣式key-value形式屬性配置(優先級高於StyleName設置的class類樣式)",
        "StyleName":"array-可選-填寫style對象有的樣式"
      },
      "Columns": { //表頭配置 必填
        "ColStyle":"object-可選-樣式key-value形式屬性配置(優先級高於StyleName設置的class類樣式)",
        "StyleName":"array-可選-填寫style對象有的樣式"
        "data": [ //表頭展示文字組 必填
          {
            "Value": "string-表頭文字",
            "ColStyle":"object-可選-樣式key-value形式屬性配置(優先級高於StyleName設置的class類樣式)",
            "StyleName":"array-可選-填寫style對象有的樣式"
          }
        ],
        //or data 為簡單的數組
        /**
         * data:["string-表頭文字"...]
         */
        "keys": ["string-表頭文字對應的數據key"] //必填
      },
      "GridData": [] //必填-列表數據項。
    }
  }
```
## <a id="styleTalk">6-樣式屬性詳解</a>
  **樣式屬性遵循pdfkit-npm.md的github開源說明文檔，在此基礎上增加一些實用屬性。**
  ### <a id="textStyle">&emsp;6.1-Text樣式</a>
  >* lineBreak-设置为false禁用所有换行
  >* align-left（默认），center，right和justify
  >* width -文本应换行的宽度（默认情况下，页面宽度减去左右边距）
  >* height -文本应剪切到的最大高度
  >* ellipsis-太长时显示在文本末尾的字符。设置为true使用默认字符。
  >* columns -文本流入的列数
  >* columnGap -每列之间的间距（默认为1/4英寸）
  >* indent -以PDF磅为单位（每磅72英寸）的缩进量
  >* paragraphGap -文本各段之间的间距
  >* lineGap -每行文字之间的间距
  >* wordSpacing -文本中每个单词之间的间距
  >* characterSpacing -文本中每个字符之间的间距
  >* fill-是否填写文字（true默认情况下）
  >* stroke -是否描边文字
  >* link -链接此文本的URL（创建注释的快捷方式）
  >* underline -是否在文字下划线
  >* strike -是否删除文字
  >* oblique-是否倾斜文字（角度或度数true）
  >* baseline-文本相对于其插入点的垂直对齐方式（值为canvas textBaseline）
  >* continued-文本段是否紧随其后。对于更改段落中间的样式很有用(不建議使用)。
  >* features- 要应用的OpenType功能标签的数组。如果未提供，则使用一组默认值。
  - isContinued -新增屬性,作用是與continued區別可以於劃線區分開(建議以此屬性為連接內容)
  - underLine -新增對象屬性，作用是可以通過對象屬性配置的形式來設置下劃線。(對劃多行空內容很有用)
    * lineWidth 線粗值
    * rowCount 行數
    * rowHeight 行高
    * X 距離第一行的開始位置靠前的距離
    * width 總長度
    * color 線顏色
    * opacity 透明度
    * lineCap 劃線類型,可設置類型如下：
    >* butt 兩邊點為正90%角的線
    >* round 兩邊點為圓弧的線
    >* square moveTo和circle搭配使用可以得到一個可設置線寬度的大小的圓。
  - padding -新增無敵可穿越周圍內容類型屬性。
    * 不受width,height設定限制,一個可以隨意跨越其他box的存在
    * 如設置Text width,那劃線終點長度將會加上paddingRight
    * 對劃方塊內容邊界很有用
    * 如設置undeLine width,劃線終點長度不會加上paddingRight
  
  ### <a id="imageStyle">&emsp;6.2-Image樣式</a>
    
  >> **此程序對pdfkit的圖片進行再次封裝<br/>**

  >> **此程序Image屬性配置：**
     >>> * url-可配置為圖片base64字符/圖片文件路徑,支持jpeg,jpg,png等格式。
     >>> * 其樣式屬性均可使用以下pdfkit-github說明

  >> **pdfkit-github文檔如下：**
     >>> * 不設置width、height-图像以全尺寸呈现
     >>> * width提供但未提供height-图像按比例缩放以适合提供的图像width
     >>> * height提供但未提供width-图像按比例缩放以适合提供的图像height
     >>> * 两者width和height提供-图像被拉伸到提供的尺寸
     >>> * scale 提供的系数-通过提供的比例系数按比例缩放图像
     >>> * fit 提供的数组-图像按比例缩放以适合传递的宽度和高度
     >>> * cover 提供的数组-图像按比例缩放以完全覆盖通过的宽度和高度定义的矩形
     >>> * link -链接此图像的URL（创建注释的快捷方式）
     >>> * goTo -定位（创建注释的快捷方式）
     >>> * destination -为此图片创建锚点  
  >> * 提供fit或cover数组时，PDFKit接受以下附加选项：
     >>> * align-水平对齐图像，可能的值是'left'，'center'和'right'
     >>> * valign-垂直对齐的图像中，可能的值是'top'，'center'和'bottom'
  
  ### <a id="checkboxStyle">&emsp;6.3-checkbox樣式</a>

  > * isContinuedNext boolean-新增屬性-是否連接下一個類型的內容，不設置默認不連接

  ### <a id="gridStyle">&emsp;6.4-grid樣式</a>
  
  > * dataRowHeight 可設置每列數據行的高度(當行數據為空時設置行高度很有用處)

## <a id="practicalTips">7-實用技巧</a>

  - isRowFollowX boolean-新增属性。应用如下：
  > * 當前一類型內容isContinued=true，此一內容類型為'checkbox'時，如想 此類型內容第二行對齊第二行設置isRowFollowX=true。
  > * 當前一類型內容isContinued=true,此-內容類型為'Grid'時,
  設置isRowFollowX=true,可以修復數據行不與表頭行對齊問題。
  
      