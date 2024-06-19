/**write by chen hai quan (jacky) 2020/03/24 ----begin*/
/**write by chen hai quan (jacky) 2020/03/31 ----第一版（完整版）end*/
/**update chen hai quan (jacky) 2020/04/01 
 * 1.多類型內容並列屬性設置增加。
 */
/**
 * update chen hai quan (jacky) 2020/04/03 
 * 1.輸出文件路徑校驗是否存在文件夾以此是否增加目標文件夾。
 * 2.增加運用async/awiat+promise 處理回調函數數據延時問題。
 * 
 * update chen hai quan 2020/08/31
 * 1.換page-y解決
 * 
 * update chen hai quan 2020/09/01
 * 1.增加圖片壓縮功能
 * 2.type-image 增加可配置quality-圖片質量百分比
 * 3.type-imageGroup 增加可配置quality-圖片質量百分比
 * 
 * update chen hai quan 2020/09/08
 * 1.增加圖片壓縮到的文件夾路徑以及imgList json文件名稱配置
 * 
 */
const PDFDocument = require('pdfkit');
const fs=require('fs');
const ImageCompress=require('images')
const path=require('path')

//===默認靜態資源路徑====begin 不同系統靜態資源路徑可能不同需改相對路徑
let filePath={
   font:'assets/fonts/',
   img:'assets/img/'
}
//字體文件路徑
let fontPaths = {
  "Deng": filePath.font+"Deng.ttf",
  "Dengb": filePath.font+"Dengb.ttf",
  "Dengl": filePath.font+"Dengl.ttf",
  "Symbola": filePath.font+"Symbola.ttf"
}
let imgPaths={
  'check':filePath.img+'check.png'
}
//===默認靜態資源路徑====end

//公共方法===begin
class pdfCommonUtil{

  async mkdirSync(filePath){
    let _filePath=filePath;
    let reduceFs=async (filePath_)=>{
      if(!fs.existsSync(path.dirname(filePath_))){
        await reduceFs(path.dirname(filePath_))
      }
      fs.mkdirSync(filePath_)
    }
    if(!fs.existsSync(_filePath)){
      await reduceFs(_filePath)
    }
  }
  async getNeedCompress(jsonPath,ImageName){
    if(!fs.existsSync(jsonPath)){
      await fs.writeFileSync(jsonPath,'{}')
    }
    const jsonDataStr=await fs.readFileSync(jsonPath,"utf8");
    if(jsonDataStr.trim() === ''){
      return true
    }
    const jsonDataParse=JSON.parse(jsonDataStr);
    return !jsonDataParse[ImageName]
  }
  
  async writeFileSync(jsonPath,ImageName){
    let jsonDataStr=await fs.readFileSync(jsonPath,'utf-8');
    if(jsonDataStr.trim() === ''){
      jsonDataStr='{}'
    }
    let jsonDataParse=JSON.parse(jsonDataStr)
    jsonDataParse[ImageName]=true;
    await fs.writeFileSync(jsonPath,JSON.stringify(jsonDataParse));
  }

  //async await 方式處理回調函數-回調在參數尾部
  async WaitFunction(paramFunc, ...args) {
    return new Promise((resolve) => {
      paramFunc(...args, (...result) => {
        resolve(result);
      });        
    });
  }
  isBase64(str){
    let exg=new RegExp('^([A-Za-z0-9+/]{4})*([A-Za-z0-9+/]{4}|[A-Za-z0-9+/]{3}=|[A-Za-z0-9+/]{2}==)$');
    return exg.test(str);
  }
  base64_decode(base64str,file){
    let bitmap = new Buffer(base64str, 'base64');
    // write buffer to file
    fs.writeFileSync(file, bitmap);
    return file
  }
  base64_encode(url){
    let buffer=fs.readFileSync(url);
    return buffer.toString('base64')
  }
  toArrayBuffer(buf) {
      var ab = new ArrayBuffer(buf.length);
      var view = new Uint8Array(ab);
      for (var i = 0; i < buf.length; ++i) {
          view[i] = buf[i];
      }
      return ab;
  }
  getStyleOptions(Style,StyleName,ColStyle) {
    let option = {};
    for (let className of StyleName) {
      if (Style[className]) {
        Object.assign(option, Style[className])
      }
    }
    Object.assign(option, ColStyle)
    return option
  }
  drawLine(doc,drawLineObj) {
    let Doc=doc;
    if(drawLineObj['lineWidth']){
      Doc.lineWidth(Number(drawLineObj.lineWidth))
      .strokeColor(drawLineObj.strokeColor)
      .strokeOpacity(Number(drawLineObj.strokeOpacity))
      .lineCap(drawLineObj.lineCap)
      .moveTo(drawLineObj.moveTo[0], drawLineObj.moveTo[1])
      .lineTo(drawLineObj.lineTo[0], drawLineObj.lineTo[1])
      .stroke();
    }
    return Doc;
  }
  drawRect(doc,drawRectObj){
    let Doc=doc;
    if(drawRectObj.lineWidth){
      Doc.lineWidth(Number(drawRectObj.lineWidth))
         .lineJoin(drawRectObj.lineJoin)
         .rect(drawRectObj.beginX,drawRectObj.beginY,drawRectObj.width,drawRectObj.height)
         .stroke()
    }
    return Doc
  }
  getBorderDrawLine(ContentBeginX,ContentBeginY,border, width, height, padding,lineGap) {
    /**
     * border 字符串設置格式 '1 #000 0.8'-'線寬度 線顏色 透明度'
     * borderAry 0-即為寬度 1-即為顏色 2-即為透明度
     */
    let borderAry = [],
        borderObj = {
          top: {},
          left: {},
          bottom: {},
          right: {}
        },
        borderItemObj = {
          lineWidth: 1,
          strokeColor: '#000',
          strokeOpacity: 1,
          lineCap: 'butt', //line cap settings
          moveTo: [0, 0],//劃線起點
          lineTo: [0, 0]
        },
        paddingObj = {
          top: padding.top ? padding.top : 0,
          bottom: padding.bottom ? padding.bottom : 0,
          right: padding.right ? padding.right : 0,
          left: padding.left ? padding.left : 0
        };
    if (Object.prototype.toString.call(border) === '[object String]') {
      borderAry = border.split(' ');
      for (let borderItem in borderObj) {
        borderItemObj.lineWidth = borderAry[0].replace('px', '');
        borderItemObj.strokeColor = borderAry[1] ? borderAry[1] : '#000';
        borderItemObj.strokeOpacity = borderAry[2] ? borderAry[2] : 1;
        borderObj[borderItem] = this.getMoveLineTo(ContentBeginX,ContentBeginY,borderItem, borderItemObj, width, paddingObj, height,lineGap);
      }
    }
    else if (Object.prototype.toString.call(border) === '[object Object]') {
      for (let borderItem2 in border) {
        borderItemObj.lineWidth = border.lineWidth ? border.lineWidth:1;
        borderItemObj.strokeColor = border.strokeColor ? border.strokeColor:'#000';
        borderItemObj.strokeOpacity = border.strokeOpacity?border.strokeOpacity:1;
        borderItemObj = this.getMoveLineTo(ContentBeginX,ContentBeginY,borderItem2, borderItemObj, width, paddingObj, height,lineGap);
        borderObj[borderItem2] = borderItemObj;
      }
    }
    else {
      throw ("border 只能為字符串或者對象類型！")
    }
    return borderObj;
  }
  getMoveLineTo(ContentBeginX,ContentBeginY,borderItem, borderItemObj, width, paddingObj, height,lineGap) {
    let newBorderItemObj = {
      lineWidth: borderItemObj.lineWidth,
      strokeColor: borderItemObj.strokeColor,
      strokeOpacity: borderItemObj.strokeOpacity,
      lineCap: borderItemObj.lineCap, //line cap settings
      moveTo: [borderItemObj.moveTo[0], borderItemObj.moveTo[0]],//劃線起點
      lineTo: [borderItemObj.lineTo[0], borderItemObj.lineTo[0]]
    };
    let lineGapLen=lineGap;//每行文字的間距
    switch (borderItem) {
      case 'top':
        newBorderItemObj.moveTo[0] = ContentBeginX - paddingObj.left;
        newBorderItemObj.moveTo[1] = ContentBeginY - paddingObj.top-3;
        newBorderItemObj.lineTo[0] = ContentBeginX + width + paddingObj.right;
        newBorderItemObj.lineTo[1] = ContentBeginY - paddingObj.top-3;
        return newBorderItemObj;
      case 'left':
        newBorderItemObj.moveTo[0] = ContentBeginX - paddingObj.left;
        newBorderItemObj.moveTo[1] = ContentBeginY - paddingObj.top-3;
        newBorderItemObj.lineTo[0] = ContentBeginX - paddingObj.left;
        newBorderItemObj.lineTo[1] = ContentBeginY + height + paddingObj.bottom-3;
        return newBorderItemObj;
      case 'bottom':
        newBorderItemObj.moveTo[0] = ContentBeginX - paddingObj.left;
        newBorderItemObj.moveTo[1] = ContentBeginY + height + paddingObj.bottom-3;
        newBorderItemObj.lineTo[0] = ContentBeginX + width + paddingObj.right;
        newBorderItemObj.lineTo[1] = ContentBeginY + height + paddingObj.bottom-3;
        return newBorderItemObj;
      case 'right':
        newBorderItemObj.moveTo[0] = ContentBeginX + width + paddingObj.right;
        newBorderItemObj.moveTo[1] = ContentBeginY - paddingObj.top-3;
        newBorderItemObj.lineTo[0] = ContentBeginX + width + paddingObj.right;
        newBorderItemObj.lineTo[1] = ContentBeginY + height + paddingObj.bottom-3;
        return newBorderItemObj;
    }
  }
}
//公共方法===end

let pdfComUtil=new pdfCommonUtil();

class Image {
  constructor(doc, {
    url = '',
    X = 0,
    Y = 0,
    quality=80,
    saveImgPath=false,
    imgJsonPath=false,
    StyleName = [],
    ColStyle={}
  }, Style, currentPosition) {
    this.doc = doc;
    this.Style = Style;
    this.currentPosition = currentPosition;
    this.url = url;
    this.X = X;
    this.Y = Y;
    this.StyleName = StyleName;
    this.ColStyle=ColStyle;
    this.quality=quality;
    this.saveImgPath=saveImgPath;
    this.imgJsonPath=imgJsonPath;
    
    this.ImageBeginX;
    this.ImageBeginY;
  }
  getX() {
    let x = this.X + this.currentPosition.x;
    return x ? x : 0;
  }
  getY() {
    let y = this.Y + this.currentPosition.y;
    return y ? y : 0;
  }
  async handlerCompress(imageWidth,imageHeight){
    let url=this.url;
    if(this.saveImgPath || this.imgJsonPath){
      await pdfComUtil.mkdirSync(this.saveImgPath)
      
      const urlSplit=url.split('/');
      const ImageName=urlSplit[urlSplit.length-1];
      const ImageNameAry=ImageName.split('.');
      const isNeedCompress=await !fs.existsSync(this.saveImgPath+'/'+ImageNameAry[0]+'_s.'+ImageNameAry[1])
      if(isNeedCompress){
        await ImageCompress(url).resize(imageWidth,imageHeight).save(this.saveImgPath+'/'+ImageNameAry[0]+'_s.'+ImageNameAry[1],{quality:this.quality})
      }
      url=this.saveImgPath+'/'+ImageNameAry[0]+'_s.'+ImageNameAry[1]
    }
    return url
  }
  async handlerImage() {

    //url 校驗
    if(pdfComUtil.isBase64(this.url)){
      this.url=pdfComUtil.base64_decode(this.url,filePath.img+'img_'+Date.now()+'.png');
    }
    if(!this.url && this.url.trim() === ''){
      throw('圖片url不能為空')
    }
    const urlExtname=path.extname(this.url);
    const imageType={'.jpg':true,'.png':true,'.jpeg':true}
    if(!imageType[urlExtname.toLocaleLowerCase()]){
      throw('只支持jpg/png格式圖片,不支持'+urlExtname+'格式文件')
    }

    this.ImageBeginX=this.getX();
    this.ImageBeginY=this.getY();
    let styleOptions=pdfComUtil.getStyleOptions(this.Style,this.StyleName,this.ColStyle);
    //設置默認寬高,為了計算位置必須要有寬高
    styleOptions['width']=styleOptions.width?styleOptions.width:80;
    styleOptions['height']=styleOptions.height?styleOptions.height:80;
    this.url="D:/Henderson/testChqDomcument/pdfImage/0acbb7b0-e842-11ea-9604-539e62ace118.jpg"
    let url = await this.handlerCompress(styleOptions['width']*2,styleOptions['height']*2)
    
    if(this.ImageBeginY>(this.doc.page.height-100-this.doc.page.margins.bottom)){
      this.doc.fillColor('#000')
      this.doc.text("第"+this.doc._pageBuffer.length+"頁",(this.doc.page.width-68),(this.doc.page.height-36), {});
      this.doc.addPage();
      this.currentPosition.y=this.doc.options.margins.top;
      this.ImageBeginY=this.currentPosition.y+this.Y;
    }
    await this.doc.image(url,this.ImageBeginX,this.ImageBeginY,styleOptions);

    //重新計算位置
    if(styleOptions.isContinued){
      this.currentPosition.x+=styleOptions.width+this.X;
      this.currentPosition.y+=this.Y; //繼承Y
    }else{
      this.currentPosition.x=this.doc.page.margins.left;
      this.currentPosition.y+=styleOptions.height+this.Y;
    }
  }
  async render() {
    await this.handlerImage()
    return { doc: this.doc, currentPosition: this.currentPosition };
  }
}
class Text {
  constructor(doc, {
    X = 0,
    Y = 0,
    StyleName = [],
    ColStyle = {},
    ImageLength=0,
    Value = ""
  }, Style, currentPosition,isCanAddPage=true) {
    this.doc = doc;
    this.Style = Style;
    this.currentPosition = currentPosition;
    this.X = X; //和之前的內容X距離，如之前沒有內容將相對於0位置
    this.Y = Y; //和之前的內容Y距離，如之前沒有內容將相對於0位置
    this.StyleName = StyleName;
    this.ColStyle = ColStyle;
    this.Value = Value;
    this.isCanAddPage=isCanAddPage;
    this.imageLength=ImageLength;
    this.ContentBeginX;
    this.ContentBeginY;
  }
  getX() {
    let x = this.X + this.currentPosition.x;
    return x ? x : 0;
  }
  getY() {
    let y = this.Y + this.currentPosition.y;
    return y ? y : 0;
  }
  handerText() {
    let styleOptions = pdfComUtil.getStyleOptions(this.Style,this.StyleName,this.ColStyle),
        text = this.Value;
        styleOptions['lineGap']=styleOptions.lineGap?styleOptions.lineGap:6;//每行文字間距
    let pageMargins = this.doc.page.margins,
        pageWidth = this.doc.page.width;
    let padding = styleOptions.padding ? styleOptions.padding:{left:0, right:0, top:0, bottom:0};
    let fontSize=styleOptions.fontSize ? styleOptions.fontSize : 14;

    this.ContentBeginX = this.getX();
    this.ContentBeginY = this.getY();
    // if(pdfComUtil.isBase64(styleOptions.font)){
    //   styleOptions.font=pdfComUtil.base64_decode(styleOptions.font,filePath.font+'font_'+Date.now()+'.ttf');
    // }
    this.doc.font(styleOptions.font ? styleOptions.font : 'Deng');
    this.doc.fontSize(fontSize);
    this.doc.fillColor(styleOptions.color ? styleOptions.color : '#000');

    if(styleOptions.border && Object.keys(styleOptions.border).length === 0){
      delete styleOptions.border
    }
    let contentWidth=Math.ceil(this.doc.widthOfString(text,styleOptions));
        // contentHeight=Math.ceil(this.doc.heightOfString(text,styleOptions));
    let contentHeight=fontSize*(Math.ceil(contentWidth/(styleOptions.width?styleOptions.width:pageWidth)))+styleOptions['lineGap']-2;
    let underlineSetWidth=0,//下劃線設置的線長
        underlineAfterHeight=0;
    let TextBoxlength=0,TextBoxHeight=0;
    let lineYDeviation=0;
    //處理下劃線===begin
      if(styleOptions.underLine){
        let underlineObj = styleOptions.underLine;
        let rowCount=underlineObj.rowCount?underlineObj.rowCount:1,
            rowHeight=underlineObj.rowHeight?underlineObj.rowHeight:fontSize+2;
        let lineToX = styleOptions.width ? styleOptions.width + this.ContentBeginX + (padding.right ? padding.right : 0)
                      :
                      ((rowCount && rowCount > 1) ? (pageWidth - pageMargins.right) : 
                        contentWidth + this.ContentBeginX + (padding.right ? padding.right : 0))
        if(underlineObj.width){
          underlineSetWidth=underlineObj.width;
          lineToX=underlineObj.width+ this.ContentBeginX;
        }
        if(rowHeight > (fontSize+2)){
          lineYDeviation=rowHeight-(fontSize+2);
          styleOptions.lineGap+=lineYDeviation-5
        }
        
        let moveToX;
        switch (underlineObj.type){
          case 'solid':
                moveToX=this.ContentBeginX;
            for(let i = 1 ; i <= rowCount ; i++ ){
              if(i>1){
                moveToX=underlineObj.X?(this.ContentBeginX-underlineObj.X):pageMargins.left;
              }
              this.doc=pdfComUtil.drawLine(this.doc,{
                lineWidth:underlineObj.height?underlineObj.height:1,
                strokeColor:underlineObj.color?underlineObj.color:'#000',
                strokeOpacity:underlineObj.opacity?underlineObj.opacity:1,
                lineCap:underlineObj.lineCap?underlineObj.lineCap:'butt',
                moveTo:[moveToX,this.ContentBeginY+rowHeight*i-lineYDeviation],
                lineTo:[lineToX,this.ContentBeginY+rowHeight*i-lineYDeviation]
              })
            }
          break;
          case 'dash':
          break;
          default:
              moveToX=this.ContentBeginX;
            for(let i = 1 ; i <= rowCount ; i++ ){
              if(i>1){
                moveToX=underlineObj.X?(this.ContentBeginX-underlineObj.X):pageMargins.left;
              }
             this.doc=pdfComUtil.drawLine(this.doc,{
                lineWidth:underlineObj.height?underlineObj.height:1,
                strokeColor:underlineObj.color?underlineObj.color:'#000',
                strokeOpacity:underlineObj.opacity?underlineObj.opacity:1,
                lineCap:underlineObj.lineCap?underlineObj.lineCap:'butt',
                moveTo:[moveToX,this.ContentBeginY+rowHeight*i-lineYDeviation],
                lineTo:[lineToX,this.ContentBeginY+rowHeight*i-lineYDeviation]
              })
            }
        }
        underlineAfterHeight=rowHeight*rowCount;
        TextBoxlength=lineToX-moveToX;
        TextBoxHeight=styleOptions.height?Number(styleOptions.height):underlineAfterHeight+styleOptions['lineGap'];
      }
    //處理下劃線===end
    
    //處理換行符 && 保留...
    text=text.replace(/\n/g,"<br/>").replace(/⋯/g,'...')
    let textAry=text.split('<br/>')
    
    if(text.trim() !== '' && this.imageLength 
      && ((this.currentPosition.y+this.doc.page.margins.bottom+150)>this.doc.page.height)){

      this.currentPosition.y=this.doc.options.margins.top;
      this.doc.fillColor('#000');
      this.doc.text("第"+this.doc._pageBuffer.length+"頁",(this.doc.page.width-68),(this.doc.page.height-36), {});
      this.doc.addPage()
      this.ContentBeginY=this.getY();

    }

    if(textAry.length>1){
      styleOptions['isContinued']=false;
      this.doc.text(textAry[0],this.ContentBeginX, this.ContentBeginY, styleOptions)
      this.currentPosition.y+=contentHeight+styleOptions.lineGap
      this.doc.text(textAry[1],this.ContentBeginX+20,this.currentPosition.y+this.Y,styleOptions)
    }else{
      this.doc.text(text, this.ContentBeginX, this.ContentBeginY, styleOptions);
    }

    //處理border===begin
      if(styleOptions.border && Object.keys(styleOptions.border).length !== 0){
        let BWidth=styleOptions.width?styleOptions.width:(underlineSetWidth?underlineSetWidth:contentWidth);
        let BHeight=styleOptions.height?styleOptions.height:(underlineAfterHeight?underlineAfterHeight:contentHeight);
        let borderDrawLine=pdfComUtil.getBorderDrawLine(this.ContentBeginX,this.ContentBeginY,styleOptions.border,BWidth,
          BHeight,padding,styleOptions.lineGap);
        let borderLine2H=0,borderLine2W=0;
        for(let BItem in borderDrawLine){
          if(borderDrawLine[BItem].lineWidth){
            this.doc=pdfComUtil.drawLine(this.doc,borderDrawLine[BItem]);
            switch (true){
              case BItem === 'top' || BItem === 'bottom':
                borderLine2H+=Number(borderDrawLine[BItem].lineWidth);
              break;
              case BItem === 'left' || BItem === 'right':
                borderLine2W+=Number(borderDrawLine[BItem].lineWidth);
              break;
            }
          }
        }
        TextBoxlength=Number(BWidth)+borderLine2W;
        TextBoxHeight=Number(BHeight)+borderLine2H;
      }
    //處理border===end
    
    //重新獲取當前位置===begin
      if(styleOptions.isContinued){
        this.currentPosition.x+=styleOptions.width?styleOptions.width+this.X:(TextBoxlength?TextBoxlength+this.X:contentWidth+this.X);
        this.currentPosition.y+=this.Y; //承接距頂設置Y
      }else{
        this.currentPosition.x=pageMargins.left;
        this.currentPosition.y+=(TextBoxHeight?TextBoxHeight:contentHeight+styleOptions.lineGap);
        let page_bottom=this.doc.page.height-(text.trim() === ''?88:8);
        if((this.currentPosition.y+16)>page_bottom){
          this.currentPosition.y=this.doc.options.margins.top;
          this.doc.fillColor('#000');
          this.doc.text("第"+this.doc._pageBuffer.length+"頁",(this.doc.page.width-68),(this.doc.page.height-36), {});
          this.doc.addPage()
        }
      }
    //重新獲取當前位置===end
  }
  render() {
    this.handerText()
    return { doc: this.doc, currentPosition: this.currentPosition };
  }
}
class Checkbox {
  constructor(doc, {
    X = 0,
    Y = 0,
    StyleName = [],
    ColStyle={},
    CheckData = []
  }, Style, currentPosition) {
    this.doc = doc;
    this.Style = Style;
    this.currentPosition = currentPosition;
    this.X = X;
    this.Y = Y;
    this.StyleName = StyleName;
    this.ColStyle = ColStyle;
    this.CheckData = CheckData;
  }
  getX() {
    let x = this.X + this.currentPosition.x;
    return x ? x : 0;
  }
  getY() {
    let y = this.Y + this.currentPosition.y;
    return y ? y : 0;
  }
  async handlerCheckBox(){
    let CheckData=this.CheckData;
    if(CheckData.length === 0){
      return
    }
    let styleOptions=pdfComUtil.getStyleOptions(this.Style,this.StyleName,this.ColStyle);
    let checkItemStyO={},TrueStyleOptions={};
    let listBeginX=this.getX();
    let checkBeginX=this.getX(),checkBeginY=this.getY();
    let checkRes,checkImagRes,TextY=0,checkAryContentLen=0;
    for(let i =0;i<CheckData.length;i++){
        checkItemStyO=pdfComUtil.getStyleOptions(this.Style,
          CheckData[i].StyleName?CheckData[i].StyleName:[],
          CheckData[i].ColStyle?CheckData[i].ColStyle:{});
        Object.assign(TrueStyleOptions,styleOptions,checkItemStyO);
        
        //默認同水平線排列
        TrueStyleOptions['isContinued']=TrueStyleOptions.isContinued?TrueStyleOptions.isContinued:true;
        TrueStyleOptions['margins']=TrueStyleOptions.margins?
            TrueStyleOptions.margins:{left:0,top:0,right:0,bottom:0};
        
        //設置默認勾選框樣式
        TrueStyleOptions['fontSize']=TrueStyleOptions.fontSize?TrueStyleOptions.fontSize:14;
        TrueStyleOptions['checkStyle']=TrueStyleOptions.checkStyle?TrueStyleOptions.checkStyle
          :{lineWidth:1,lineJoin:'miter',width:12,height:12}
        TextY=-(TrueStyleOptions.fontSize-TrueStyleOptions.checkStyle.height)/2;
          if(TrueStyleOptions.isContinued){
            checkAryContentLen+=(TrueStyleOptions.margins.left?TrueStyleOptions.margins.left:0)+
            TrueStyleOptions.checkStyle.width+this.doc.widthOfString(CheckData[i].value,TrueStyleOptions);
          }
        if(i>0){
          //處理margins margins為第二個開始每個checkbox相對位置
          if(TrueStyleOptions.isContinued){
            if(checkAryContentLen > (this.doc.page.width-this.doc.page.margins.left-this.doc.page.margins.right-listBeginX)){
              checkBeginX=TrueStyleOptions.isRowFollowX?listBeginX:this.doc.page.margins.left+this.X;
              checkBeginY+=TrueStyleOptions.checkStyle.height+
              (TrueStyleOptions.margins.bottom?TrueStyleOptions.margins.bottom:0)+6;
              
              this.currentPosition.y=checkBeginY-3;
              checkAryContentLen=(TrueStyleOptions.margins.left?TrueStyleOptions.margins.left:0)+3+
              TrueStyleOptions.checkStyle.width+this.doc.widthOfString(CheckData[i].value,TrueStyleOptions)
            }else{
              checkBeginX+=(TrueStyleOptions.margins.left?TrueStyleOptions.margins.left:0)+3+
                TrueStyleOptions.checkStyle.width+this.doc.widthOfString(CheckData[i-1].value,TrueStyleOptions);
              
              checkBeginY=this.currentPosition.y;
            }
          }else{
            checkBeginX=this.getX();
            checkBeginY+=TrueStyleOptions.checkStyle.height+
              (TrueStyleOptions.margins.bottom?TrueStyleOptions.margins.bottom:0);
          }
        }
        //1-劃check框 || 如勾選加上勾選圖片
        TrueStyleOptions.checkStyle['beginX']=checkBeginX;
        TrueStyleOptions.checkStyle['beginY']=checkBeginY;
        this.doc=pdfComUtil.drawRect(this.doc,TrueStyleOptions.checkStyle)
        this.currentPosition.x=checkBeginX+TrueStyleOptions.checkStyle.width+3;
        if(CheckData[i].isChecked){
          checkImagRes=await new Image(this.doc,{
            url:imgPaths.check,
            ColStyle:{
              width:TrueStyleOptions.checkStyle.width-TrueStyleOptions.checkStyle.lineWidth*2,
              height:TrueStyleOptions.checkStyle.height-TrueStyleOptions.checkStyle.lineWidth*2
            }
          },this.Style,{
            x:checkBeginX+TrueStyleOptions.checkStyle.lineWidth,
            y:checkBeginY+TrueStyleOptions.checkStyle.lineWidth
          }).render();
          this.doc=checkImagRes.doc;
        }

        //2-渲染文字
        if(i === (CheckData.length-1)){
          TrueStyleOptions["isContinued"]=styleOptions.isContinuedNext?
            styleOptions.isContinuedNext:false;
        }
        checkRes=new Text(this.doc,{
          X:0,
          Y:0,
          StyleName:[],
          ColStyle:TrueStyleOptions,
          Value:CheckData[i].value
        },this.Style,this.currentPosition).render();
        this.doc=checkRes.doc;
        this.currentPosition=checkRes.currentPosition;
    }
  }
  render() {
    this.handlerCheckBox();
    return { doc: this.doc, currentPosition: this.currentPosition };;
  }
}
class Grid {
  constructor(doc, {
    X = 0,
    Y = 0,
    GridWidth=0,
    Title = {}, //繼承Text
    Columns = {
      ColStyle:{},
      StyleName:[],
      data:[], //每個對象繼承Text
      keys:[]  //每個欄位的key
    },
    GridData = []
  }, Style, currentPosition) {
    this.doc = doc;
    this.Style = Style;
    this.currentPosition = currentPosition;
    this.X = X;
    this.Y = Y;
    this.pageWidth=this.doc.page.width;
    this.pageMargins=this.doc.page.margins;
    this.GridWidth=GridWidth?GridWidth:this.pageWidth-(this.pageMargins.left+this.pageMargins.right)-this.X;
    this.Title = Title; //繼承text
    this.columns = Columns;
    this.gridData = GridData;
    this.gridToalHeight=0;

    this.columnStyleOptions=[];
    this.ColStyleBorderLeft={};

    this.columnBeginX=0;
  }

  handlerGrid(){
    this.handlerTitle();
    this.handlerColumns();
    this.handlerGridData();
  }
  handlerGridData(){
    
    let gridData=this.gridData;
    let columnStyleOptions=this.columnStyleOptions;
    let keys=this.columns.keys;
    if(gridData.length === 0){
      return;
    }
    
    let gridDataValue='';
    let gridDataHeight=0,lineGapHeight=0,valueLineGapHeight=0;
    let valueContentLen=0,valueContentHeight=0;
    //第一次循環計算出內容高度 && 行總距
    for(let item of gridData){
      for(let i = 0;i<columnStyleOptions.length;i++){
        gridDataValue=item[keys[i]]?item[keys[i]]:'';
        valueContentLen=this.doc.widthOfString(gridDataValue,columnStyleOptions[i]);
        valueContentHeight=this.doc.heightOfString(gridDataValue,columnStyleOptions[i]);
        gridDataHeight=gridDataHeight>valueContentHeight?gridDataHeight:valueContentHeight;
        valueLineGapHeight=columnStyleOptions[i].lineGap*(Math.ceil(valueContentLen/columnStyleOptions[i].width)+1);
        lineGapHeight=lineGapHeight>valueLineGapHeight?lineGapHeight:valueLineGapHeight; 
      }
    }

    //第二次循環劃數據
    let gridDataX=0,gridDataY=0;
    let gridDataRes,isNextRow=false,gridDataToalHeight=0;
    const columnAfterPosition={
      x:this.currentPosition.x,
      y:this.currentPosition.y
    }
    for(let k=0;k<gridData.length;k++){
        let item2=gridData[k];
        isNextRow=true;
      for(let j = 0;j<columnStyleOptions.length;j++){
        gridDataValue=item2[keys[j]]?item2[keys[j]]:'';
        columnStyleOptions[j]['height']=columnStyleOptions[j].dataRowHeight?columnStyleOptions[j].dataRowHeight:
          gridDataHeight+(columnStyleOptions[j].padding.top?columnStyleOptions[j].padding.top:0)+
          (columnStyleOptions[j].padding.bottom?columnStyleOptions[j].padding.bottom:0)+3;
        columnStyleOptions[j].padding['top']=2;
        if(isNextRow){
          gridDataToalHeight+=columnStyleOptions[j]['height']+
            columnStyleOptions[j].padding.top+
            (columnStyleOptions[j].padding.bottom?columnStyleOptions[j].padding.bottom:0);
          isNextRow=false;
        }

        delete columnStyleOptions[j].border.top;
     
        if(j === 0){
          gridDataX=columnStyleOptions[j].isRowFollowX?this.columnBeginX:this.X;
          gridDataY=0;
          if(columnStyleOptions.length === 1){
            gridDataY=this.Y;
          }
        }else{
          if(columnStyleOptions.length>7){
            if(j<Math.ceil(columnStyleOptions.length/2)){
              gridDataX=j === 1?0:-((0.36-(0.16/70)*columnStyleOptions.length)*j);
            }else{
              gridDataX=-(.1*j);
            }
          }else{
            gridDataX=j === 1?0:-(0.26*j);
          }
          gridDataY=0;
        }
        if(k === (gridData.length-1) && j === (columnStyleOptions.length-1)){
          columnStyleOptions[j]['isContinued']=columnStyleOptions[j].isContinuedNext?
            columnStyleOptions[j].isContinuedNext:false;
        }
        gridDataRes=new Text(this.doc,
          {X:gridDataX,Y:gridDataY,StyleName:[],ColStyle:columnStyleOptions[j],Value:gridDataValue},
          this.Style,this.currentPosition).render();
        this.doc=gridDataRes.doc;
        this.currentPosition=gridDataRes.currentPosition;
        if(k === (gridData.length-1) && j === (columnStyleOptions.length-1) && columnStyleOptions[j]['isContinued']){
          this.gridToalHeight+=gridDataToalHeight;
          this.currentPosition.y-=(this.gridToalHeight-gridDataToalHeight+columnStyleOptions[j]['height']/2-3)
        }
      }
    }
    if(columnStyleOptions.length === 1){
      this.currentPosition.y+=this.Y
      return;
    }
    
    //補充劃left線
    let leftLineToY=columnAfterPosition.y-12+gridDataToalHeight;
    if(!this.Title.Value || this.Title.Value.trim() === ''){
      leftLineToY=columnAfterPosition.y+gridDataToalHeight-5;
    }
    let leftBeginX=columnStyleOptions[0].isRowFollowX?
        this.columnBeginX+columnStyleOptions[0].width-this.X-8:
        this.X+columnAfterPosition.x-(columnStyleOptions[0].padding.left?columnStyleOptions[0].padding.left:0);
    pdfComUtil.drawLine(this.doc,{
      lineWidth:Number(this.ColStyleBorderLeft.lineWidth),
      strokeColor:this.ColStyleBorderLeft.strokeColor?this.ColStyleBorderLeft.strokeColor:'#000',
      strokeOpacity:Number(this.ColStyleBorderLeft.strokeOpacity)?Number(this.ColStyleBorderLeft.strokeOpacity):1,
      lineCap:this.ColStyleBorderLeft.lineCap?this.ColStyleBorderLeft.lineCap:'butt',
      moveTo:[
        leftBeginX,
        columnAfterPosition.y-5
      ],
      lineTo:[
        leftBeginX,
        leftLineToY
      ]
    })
  }
  handlerColumns(){
    let columns=this.columns;
    if(columns.data.length === 0){
      return;
    }
    let ColStyle=columns.ColStyle?columns.ColStyle:{},
        StyleName=columns.StyleName?columns.StyleName:[];
     //設置默認邊框線
      let ColStyleBorderType=Object.prototype.toString.call(ColStyle.border);
      let ColStyleBorderAry=[],ColStyleBorderItemObj={};
      this.ColStyleBorderLeft={}; //表頭第一格需要劃border left，其他格均不需要,另外處理。
      switch (ColStyleBorderType){
        case '[object String]':
          ColStyleBorderAry=ColStyle.border.split(' ');
          ColStyleBorderItemObj={
            lineWidth:ColStyleBorderAry[0],
            strokeColor:ColStyleBorderAry[1],
            strokeOpacity:ColStyleBorderAry[2]
          }
          ColStyle['border']={
            top:ColStyleBorderItemObj,
            bottom:ColStyleBorderItemObj,
            right:ColStyleBorderItemObj
          }
          this.ColStyleBorderLeft=ColStyleBorderItemObj
          break;
        case '[object Object]':
          ColStyle['border']={
            top:ColStyle.border.top?ColStyle.border.top:{lineWidth:1},
            bottom:ColStyle.border.bottom?ColStyle.border.bottom:{lineWidth:1},
            right:ColStyle.border.right?ColStyle.border.right:{lineWidth:1}
          }
          this.ColStyleBorderLeft=ColStyle.border.left?ColStyle.border.left:{lineWidth:1}
          break;
        default:
          ColStyle['border']={top:{lineWidth:1},bottom:{lineWidth:1},right:{lineWidth:1}}
          this.ColStyleBorderLeft={lineWidth:1};
      }
     //設置默認width 默認等均寬度
     ColStyle['width']=ColStyle.width?ColStyle.width:(this.GridWidth/columns.data.length);
     //設置默認文字居中
     ColStyle['align']=ColStyle.align?ColStyle.align:'center';
     //styeOptions Grid公共樣式
    let styleOptions={};
    let columnDataItemType;
    let columnRes,columnValue='',columnX=0,colunmY=0;
    let columnHeight=0,lineGapHeight=0;
    let valueContentLen=0;
    //第一次循環計算表內容的高度 && 行間距總高
    for(let i = 0;i<columns.data.length;i++){
      styleOptions=pdfComUtil.getStyleOptions(this.Style,StyleName,ColStyle);
      columnDataItemType=Object.prototype.toString.call(columns.data[i]);
      columnValue=columns.data[i];
      if(columnDataItemType === '[object Object]'){
        Object.assign(styleOptions,pdfComUtil.getStyleOptions(this.Style,
          columns.data[i].StyleName?columns.data[i].StyleName:[],
          columns.data[i].ColStyle?columns.data[i].ColStyle:{}))
        columnValue=columns.data[i].Value;
      }
      styleOptions['isContinued']=(i !== columns.data.length-1);
      styleOptions['lineGap']=styleOptions.lineGap?styleOptions.lineGap:6;

      valueContentLen=this.doc.widthOfString(columnValue,styleOptions);
      columnHeight=columnHeight>this.doc.heightOfString(columnValue,styleOptions)?
        columnHeight:this.doc.heightOfString(columnValue,styleOptions);
      lineGapHeight=lineGapHeight>styleOptions.lineGap*(Math.ceil(valueContentLen/styleOptions.width)+1)?
        lineGapHeight:styleOptions.lineGap*(Math.ceil(valueContentLen/styleOptions.width)) 
    }
    //第二次循環賦值內容高度並劃表頭
    this.columnStyleOptions=[];
    for(let j=0;j<columns.data.length;j++){
      styleOptions=pdfComUtil.getStyleOptions(this.Style,StyleName,ColStyle);
      columnDataItemType=Object.prototype.toString.call(columns.data[j]);
      columnValue=columns.data[j];
      if(columnDataItemType === '[object Object]'){
        Object.assign(styleOptions,pdfComUtil.getStyleOptions(this.Style,
          columns.data[j].StyleName?columns.data[j].StyleName:[],
          columns.data[j].ColStyle?columns.data[j].ColStyle:{}));
        columnValue=columns.data[j].Value;
      }
      styleOptions['isContinued']=(j !== columns.data.length-1);
      styleOptions['padding']=styleOptions.padding?styleOptions.padding
            :{top:0,left:0,bottom:0,right:0};
      styleOptions['lineGap']=styleOptions.lineGap?styleOptions.lineGap:3;
      
      styleOptions['height']=styleOptions.height?styleOptions.height
            :columnHeight-this.X*0.21+(styleOptions.padding.top?styleOptions.padding.top:0)+
              (styleOptions.padding.bottom?styleOptions.padding.bottom:0)
              -lineGapHeight+5;
      this.gridToalHeight=styleOptions.height;
      if(j === 0){
        columnX=this.X;
        this.columnBeginX=this.X+this.currentPosition.x-this.doc.page.margins.left;
        colunmY=this.Title.Value?6:this.Y;
        styleOptions.border['left']=this.ColStyleBorderLeft;
      }else{
        columnX=-1;
        colunmY=0;
        if(styleOptions.border.left){
          delete styleOptions.border.left;
        }
      }
      this.columnStyleOptions.push(styleOptions);
      columnRes=new Text(this.doc,
        {X:columnX,Y:colunmY,StyleName:[],ColStyle:styleOptions,Value:columnValue},
        this.Style,this.currentPosition).render();
      this.doc=columnRes.doc;
      this.currentPosition=columnRes.currentPosition;
    }
  }
  handlerTitle(){
    let TitleObj=this.Title;
    if(!TitleObj.Value || TitleObj.Value.trim() === ''){
      return
    }
    
    TitleObj['X']=TitleObj.X?TitleObj.X+this.X:this.X;
    TitleObj['Y']=TitleObj.Y?TitleObj.Y+this.Y+9:this.Y+9;

    if(TitleObj.ColStyle){
      //標題默認居中
      TitleObj.ColStyle['align']=TitleObj.ColStyle.align?TitleObj.ColStyle.align:"center";
      //繼承GridWidth
      TitleObj.ColStyle['width']=TitleObj.ColStyle.width?TitleObj.ColStyle.width:this.GridWidth;
    }else{
      TitleObj['ColStyle']={
        align:"center",
        width:this.GridWidth
      }
    }
    let res=new Text(this.doc,TitleObj,this.Style,this.currentPosition).render();
    this.doc=res.doc;
    this.currentPosition=res.currentPosition;
  }
  
  render() {
    this.handlerGrid();
    return { doc: this.doc, currentPosition: this.currentPosition };
  }
}
class Lable{
  constructor(doc, {
    X = 0,
    Y = 0,
    StyleName = [],
    ColStyle = {},
    Value = ""
  }, Style, currentPosition) {
    this.doc = doc;
    this.Style = Style;
    this.currentPosition = currentPosition;
    this.X = X; //和之前的內容X距離，如之前沒有內容將相對於0位置
    this.Y = Y; //和之前的內容Y距離，如之前沒有內容將相對於0位置
    this.StyleName = StyleName;
    this.ColStyle = ColStyle;
    this.Value = Value;

    this.ContentBeginX;
    this.ContentBeginY;
  }
  getX() {
    let x = this.X + this.currentPosition.x;
    return x ? x : 0;
  }
  getY() {
    let y = this.Y + this.currentPosition.y;
    return y ? y : 0;
  }
  handerText() {
    let styleOptions = pdfComUtil.getStyleOptions(this.Style,this.StyleName,this.ColStyle),
        text = this.Value;
        styleOptions['lineGap']=styleOptions.lineGap?styleOptions.lineGap:6;;//每行文字間距
    
    let pageMargins = this.doc.page.margins;
    let fontSize=styleOptions.fontSize ? styleOptions.fontSize : 14;

    this.ContentBeginX = this.getX();
    this.ContentBeginY = this.getY();
    this.doc.font(styleOptions.font ? styleOptions.font : 'Deng');
    this.doc.fontSize(fontSize);
    this.doc.fillColor(styleOptions.color ? styleOptions.color : '#000');
    
    this.doc.text(text, this.ContentBeginX, this.ContentBeginY, styleOptions);

    let contentWidth=Math.ceil(this.doc.widthOfString(text,styleOptions)),
        contentHeight=Math.ceil(this.doc.heightOfString(text,styleOptions));

    let TextBoxlength=0,TextBoxHeight=0;
   
    //重新獲取當前位置===begin
      if(styleOptions.isContinued){
        this.currentPosition.x+=(TextBoxlength?TextBoxlength+this.X:contentWidth+this.X);
        this.currentPosition.y+=this.Y; //承接距頂設置Y
      }else{
        this.currentPosition.x=pageMargins.left;
        this.currentPosition.y+=(TextBoxHeight?TextBoxHeight:contentHeight+styleOptions.lineGap);
      }
    //重新獲取當前位置===end
  }
  render() {
    this.handerText()
    return { doc: this.doc, currentPosition: this.currentPosition };
  }
}

class ImageGroup{
  constructor(doc,data,styleClass,currentPosition){
    this.doc=doc;
    this.listData=data;
    this.styleClass=styleClass;
    this.currentPosition=currentPosition;
    this.jsonData=[]
  }
  groupHeaderLine(top,isLast){
    return{
      Type:'Text',
      Y:0,
      Obj:{
        "Value":" ",
        "ColStyle":{
            "width":570,
            "height":1,
            "isContinued":isLast,
            "border":{
              "top":"1 #000 1"
            },
            "padding":{
              "top":top
            }
        }
      }
    }
  }
  groupVerticalLine(height){
    return{
      Type:'Text',
      X:100,
      Y:-3,
      Obj:{
        "Value":" | ",
        "ColStyle":{
            "width":1,
            "height":1,
            "border":{
              "left":"1 #000 1"
            },
            "padding":{
              "top":height*2+3
            }
        }
      }
    }
  }
  ImageDataMode(url,X,Y,link,imageWidth,imageHeight,quality,saveImgPath,imgJsonPath){
    return {
      "Type":"Image",
      "Obj":{
          "url":url,
          "X":X,
          "Y":Y,
          quality:quality || 50,
          saveImgPath: saveImgPath || false,
          imgJsonPath: imgJsonPath || false,
          "ColStyle":{
            "width":imageWidth || 80,
            "height":imageHeight || 80,
            "link":link || '',
            "border":{
              "top":"1 #000 1"
            },
            "padding":{
              "top":8
            }
          }
      }
    }
  }
  groupNameDataMode(value,Y,topBottom,width,height,isNeedTop,fontSize,ImageLength){
    let obj={
      "Type":"Text",
      "Obj":{
        "Value":value,
        "Y":Y,
        ImageLength:ImageLength || 0,
        "ColStyle":{
            "isContinued":false,
            "fontSize":fontSize || 14,
            "border":{
              // "left":"1 #000 1",
              // "right":"1 #000 1",
              // "bottom":"1 #000 1"
            },
            "padding":{
              "left":8,
              "right":8,
              // "top":topBottom,
              // "bottom":topBottom
            }
        }
      }
    }
    // if(isNeedTop){
    //   obj.Obj.ColStyle.border['top']="1 #000 1";
    // }
    return obj
  }
  TextDataMode(value,X,width,isContinued,link){
    let obj={
      "Type":"Text",
      "Obj":{
          "Value":value,
          "X":X,
          "ColStyle":{
              "width":width,
              "align":"center",
              "color":"blue",
              "link":link,
              "isContinued":isContinued,
              // "border":{
              //     "bottom":"1 #000 1"
              // },
              "padding":{
                  "bottom":5
              }
          }
      }
    }
    if(value.trim() !== ''){
      obj.Obj.ColStyle['underline']={
        color:"blue"
      } 
    }else{
      obj.Obj.ColStyle['padding']={
        bottom:0
      }
    }
    return obj
  }

  async handlerImageGroup(){
     let listData=this.listData,jsonData=this.jsonData;
     let groupJson=[],allHeight=0;
     for(let i=0;i<listData.length;i++){
        let groupHeight=0,
          groupNamePaddingTopB=0,
          groupWidth=listData[i].groupNameStyle.width || 45,
          groupFontSize=listData[i].groupNameStyle.fontSize || 14;

        let imgsJsonData=[],itemIndex=-1,imgRow=0;
        //每組圖片數據====begin
        if(listData[i].imgsData){
              imgsJsonData=[];
          let imgsData=listData[i].imgsData;
          let imgAry=imgsData.imgAry || [],
              itemX=imgsData.itemX || 8,
              itemY=imgsData.itemY || 8,
              quality=imgsData.quality || 50,
              saveImgPath=imgsData.saveImgPath || false,
              imgJsonPath=imgsData.imgJsonPath || false;
          let imgWidth=imgsData.width || 80,
              imgHeight=imgsData.height || 80;
          for(let j=0;j<imgAry.length;j++){

             groupHeight=imgHeight+28;
             groupNamePaddingTopB=(groupHeight-(itemY*2))/2;

            //  if(j === 0){
            //    itemX+=5
            //  }
             itemIndex++;
            //  let y=(itemIndex === 0)?3+((imgRow && imgAry[j].title.trim() === '')?-19:0):-imgHeight,
            //      imgLink=imgAry[j].link || '';
             let y=(itemIndex === 0)?0+((imgRow && imgAry[j].title.trim() === '')?0:0):-imgHeight,
                 imgLink=imgAry[j].link || '';

            //  let textX=22+100*itemIndex;
             let textX = (itemX*(itemIndex+1)) + (imgWidth*itemIndex);
             let isContinued=true;
             
             let iPageWidth = this.doc.page.width - 2*itemX - this.doc.page.margins.right -imgWidth;

             if((textX+imgWidth+itemX) > iPageWidth){
               imgRow++;
               isContinued=false;
               itemIndex=-1;
             }
             if(j === (imgAry.length-1)){isContinued=false}
             imgsJsonData=imgsJsonData.concat([
               this.ImageDataMode(imgAry[j].url,itemX,y,imgLink,imgWidth,imgHeight,quality,saveImgPath,imgJsonPath),
               this.TextDataMode(imgAry[j].title.trim()+' ',textX,imgWidth,isContinued,imgLink)
             ])
          }
        }
        //每組圖片數據====end

        let groupName=listData[i].groupName || ' ';
        let contentHeight=this.doc.heightOfString(groupName,{width:groupWidth,height:groupHeight,fontSize:groupFontSize})
        groupNamePaddingTopB=(groupHeight-contentHeight)/2;
        let groupY=(i===0)?0:-3,groupHeaderTop=(i===0)?-2:1;
        let groupNameData=[this.groupNameDataMode(groupName,groupY,groupNamePaddingTopB,
            groupWidth,groupHeight,i===0,groupFontSize,imgsJsonData.length)]
        
        groupJson=groupJson.concat(groupNameData,this.groupHeaderLine(groupHeaderTop,false),imgsJsonData)
        if(imgsJsonData.length === 0){
          groupJson.push({
            "Type":"Text",
            "Obj":{
              "Value":' ',
              "Y":80
            }
          })
        }
        allHeight+=groupHeight;
     }
     this.jsonData=jsonData.concat(groupJson)
  }
  render(){
    this.handlerImageGroup()
    return this.jsonData
  }
}

class hsPDFUtil {
  constructor(
    {
      outFileName = 'dome_' + Date.now() + '.pdf',
      Author,
      Subject,
      Keywords,
      PageSize="LETTER",
      Space,
      font,
      Style = {},
      Title = [],
      Header = [],
      Body = [],
      Footer = []
    }
  ) {
    this.FileName = outFileName;
    this.Author = Author;
    this.Subject = Subject;
    this.Keywords = Keywords;
    this.Space = Space;
    this.SpaceDefault = {
      top: 0,
      bottom: 0,
      left: 8,
      right: 8
    }
    this.font = font;
    this.fontDefault = ["Deng", "Dengb", "Dengl", "Symbola"];
    this.Style = Style;
    this.Title = Title;
    this.Header = Header;
    this.Body = Body;
    this.Footer = Footer;
    this.PageSize=PageSize;
    this.currentPosition={x:0,y:0};
    this.doc;
  }
  initPDF(callback) {
    this.doc = new PDFDocument({
      bufferPages: true,
      Title: this.FileName,
      Author: this.Author ? this.Author : '',
      Subject: this.Subject ? this.Subject : '',
      Keywords: this.Keywords ? this.Keywords : '',
      margins: this.Space ? this.Space : this.SpaceDefault,
      size:this.PageSize
    })
    if (this.doc) {
      //註冊字體
      let fonts = this.font ? this.font : this.fontDefault;
      for (let font of fonts) {
        if (fontPaths[font]) {
          this.doc.registerFont(font, fontPaths[font])
        }
      }
      //設置內容起點
      this.currentPosition.x=this.doc.page.margins.left;
      this.currentPosition.y=this.doc.page.margins.top;

      typeof callback === 'function' && callback()
    }

  }
  async processor() { //分配處理
    let structure = ["Title", "Header", "Body", "Footer"];
    let res;
    /**
     * this[nodeKey] 即為this.Title || this.Header ...
     */
    for (let nodeKey of structure) {
      for (let item of this[nodeKey]) {
        switch (item.Type) {
          case "Text":
            res = await new Text(this.doc, item.Obj, this.Style, this.currentPosition,item.isCanAddPage).render();
            this.doc = res.doc;
            this.currentPosition = res.currentPosition;
            break;
          case "Lable":
            res = await new Lable(this.doc, item.Obj, this.Style, this.currentPosition).render();
            this.doc = res.doc;
            this.currentPosition = res.currentPosition;
            break;
          case "Image":
            res = await new Image(this.doc, item.Obj, this.Style, this.currentPosition).render();
            this.doc = res.doc;
            this.currentPosition = res.currentPosition;
            break;
          case "ImageGroup":
            this.Body = await new ImageGroup(this.doc,item.data,this.Style,this.currentPosition).render();
            await this.processor()
            break;
          case "Checkbox":
            res = await new Checkbox(this.doc, item.Obj, this.Style, this.currentPosition).render();
            this.doc = res.doc;
            this.currentPosition = res.currentPosition;
            break;
          case "Grid":
            res = await new Grid(this.doc, item.Obj, this.Style, this.currentPosition).render();
            this.doc = res.doc;
            this.currentPosition = res.currentPosition;
            break;
          default:
            throw ('無' + item.Type + "類型！")
        }
      }
    }

    this.doc.fillColor('#000');
    this.doc.text("第"+this.doc._pageBuffer.length+"頁",(this.doc.page.width-68),(this.doc.page.height-36), {});

  }
  async render(type){
    try {
      console.time('pdfUtilTime:')
      await this.initPDF()
      await this.processor()
      console.timeEnd('pdfUtilTime:')
      let fs;
      let FilePath=this.FileName.substring(0,this.FileName.lastIndexOf('/')+1);
      switch (type) {
        case "file":
          fs =require('fs');
          await pdfComUtil.mkdirSync(FilePath)
          this.doc.pipe(fs.createWriteStream(this.FileName))
          this.doc.end()
          return {
            status: 1,
            data: { fileName: this.FileName }
          }
        case "stream":
          let stream = this.doc.pipe(require('blob-stream')());
          this.doc.end();
          let getStream= (stream,cb)=>{
            stream.on('finish',function(){
              cb(stream);
            })
          }
          let [res]= await pdfComUtil.WaitFunction(getStream,stream);
          return {
            status: 1,
            data: res
          }
        default:
          fs =require('fs');
          await pdfComUtil.mkdirSync(FilePath)
          this.doc.pipe(fs.createWriteStream(this.FileName))
          this.doc.end()
          return {
            status: 1,
            data: { fileName: this.FileName }
          }
      }
    } catch (err) {
      return {
        status: 0,
        data: err
      }
    }
  }
}

module.exports = hsPDFUtil

/**
 * 註-1
 * pdfkit自帶字體
 *  'Courier'
    'Courier-Bold'
    'Courier-Oblique'
    'Courier-BoldOblique'
    'Helvetica'
    'Helvetica-Bold'
    'Helvetica-Oblique'
    'Helvetica-BoldOblique'
    'Symbol'
    'Times-Roman'
    'Times-Bold'
    'Times-Italic'
    'Times-BoldItalic'
    'ZapfDingbats'
 */

/**
 * 註-2
 * pdfkit Text樣式
 *
 * lineBreak-设置为false禁用所有换行
 * align-left（默认），center，right和justify
   width -文本应换行的宽度（默认情况下，页面宽度减去左右边距）
   height -文本应剪切到的最大高度
   ellipsis-太长时显示在文本末尾的字符。设置为true使用默认字符。
   columns -文本流入的列数
   columnGap -每列之间的间距（默认为1/4英寸）
   indent -以PDF磅为单位（每英寸72英寸）的缩进量
   paragraphGap -文本各段之间的间距
   lineGap -每行文字之间的间距
   wordSpacing -文本中每个单词之间的间距
   characterSpacing -文本中每个字符之间的间距
   fill-是否填写文字（true默认情况下）
   stroke -是否描边文字
   link -链接此文本的URL（创建注释的快捷方式）
   underline -是否在文字下划线
   strike -是否删除文字
   oblique-是否倾斜文字（角度或度数true）
   baseline-文本相对于其插入点的垂直对齐方式（值为canvas textBaseline）
   continued-文本段是否紧随其后。对于更改段落中间的样式很有用。
   features- 要应用的OpenType功能标签的数组。如果未提供，则使用一组默认值。
   isContinued -本人新增屬性,作用是可以於劃線區分開
 */

/**
 * 註-3
 * 劃線 lineCap參數說明
 * butt 兩邊點為正90%角的線
 * round 兩邊點為圓弧的線
 * square moveTo和circle搭配使用可以得到一個可設置線寬度的大小的圓。
 */

 /**
  * 註-4
  * 新增 padding 屬性說明
  * 1.不受width,height設定限制,一個可以隨意跨越其他box的存在。
  * 2.如設置Text width,那劃線終點長度將會加上paddingRight
  * 3.對劃方塊內容邊界很有用
  * 4.如設置undeLine width,劃線終點長度不會加上paddingRight
  */

  /**
   * 註-5 pdfKit 圖像
   *  width、height都未提供-图像以全尺寸呈现
      width提供但未提供height-图像按比例缩放以适合提供的图像width
      height提供但未提供width-图像按比例缩放以适合提供的图像height
      两者width和height提供-图像被拉伸到提供的尺寸
      scale 提供的系数-通过提供的比例系数按比例缩放图像
      fit 提供的数组-图像按比例缩放以适合传递的宽度和高度
      cover 提供的数组-图像按比例缩放以完全覆盖通过的宽度和高度定义的矩形

      提供fit或cover数组时，PDFKit接受以下附加选项：
        align-水平对齐图像，可能的值是'left'，'center'和'right'
        valign-垂直对齐的图像中，可能的值是'top'，'center'和'bottom'
   */
  /**
   * 註-6 underLine 新增對象
   * lineWidth
   * rowCount 行數
   * rowHeight 行高
   * X 距離第一行的開始的靠前的距離
   * width 總長度
   * color 線顏色
   * opacity 透明度
   * lineCap 劃線類型 詳情 註-3
   */

   /**
    * 註-7 grid
    * dataRowHeight 可設置每列數據行的高度(當行數據為空時設置行高度很有用處)
    */

  /**
   * 註-8 Checkbox
   * isContinuedNext false/true 是否連接下一個類型的內容 不設置默認不連接
   */

  /**
   * 註-9 isRowFollowX
   * 當前一類型內容isContinued=true
   * 1.此一內容類型為'checkbox'時
   * 如想 此類型內容第二行對齊第二行設置isRowFollowX=true
   * 2.此-內容類型為'Grid'時
   * 請設置isRowFollowX=true,可以修復數據行不與表頭行對齊問題。
   */