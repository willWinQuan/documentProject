
const { hldFileBo } =require('../bo');
const {FilePath} =require('../../config');
const images =require("images")
const path=require('path')

const fs=require('fs')
const image=require('imageinfo');

class fileController {
   
   static async getFile(ctx){
      const param =ctx.request.body;
      //取token裡的文件類型 && 返回的數據類型
      /**
       * fileType-docx/pdf/excel
       * dataType-buff/file
       * docx&&excel 只可返回buff/file
       * pdf 只可返回file(暫時)
       */
      if(!param.fileType || !param.dataType){
         ctx.body={satuts:0,msg:"文件類型或返回的數據類型不能為空！"}
         return;
      }
      const fileType=param.fileType,dataType=param.dataType;
      let fileSuffix='';
      switch (fileType){
         case 'docx':
            fileSuffix=".docx"
         break;
         case 'pdf':
            fileSuffix=".pdf"
         break;
         case 'excel':
           fileSuffix=".xlsx"
         break;
         case 'docxToPdf':
            fileSuffix=".pdf"
         break;
      }
      const opt={
        fileType:fileType,
        dataType:dataType,
        inputFileName:param.inputFileName?param.inputFileName:'',
        outFileName:FilePath+'documention/outFile_'+Date.now()+fileSuffix,
        outFilePath:FilePath+'documention',
        data:param.data
      } 
      
      let res=await hldFileBo.getFile(opt);
      ctx.body=res;
   }
   static async testEditPDF(ctx){
      try{
         let param =ctx.request.body;

         // //1-pdf2json 加载有标记的模板pdf,得到标记符的x,y text,and option,替换标记text


         // const pdfjsLib = require("pdfjs-dist/legacy/build/pdf.js");
         // const fs = require("fs")
         // const pdfBuffer=fs.readFileSync('D:/Ricardo/documentationProject/my-project/doc/Reporaaaaaat1234.pdf','binary');
         // const loadingTask = pdfjsLib.getDocument(pdfBuffer);
         // loadingTask.promise
         // .then(function (doc) {
         //    const numPages = doc.numPages;
         //    console.log("# Document Loaded");
         //    console.log("Number of Pages: " + numPages);
         //    console.log();

         //    let lastPromise; // will be used to chain promises
         //    lastPromise = doc.getMetadata().then(function (data) {
         //       console.log("# Metadata Is Loaded");
         //       console.log("## Info");
         //       console.log(JSON.stringify(data.info, null, 2));
         //       console.log();
         //       if (data.metadata) {
         //       console.log("## Metadata");
         //       console.log(JSON.stringify(data.metadata.getAll(), null, 2));
         //       console.log();
         //       }
         //    });

         //    const loadPage = function (pageNum) {
         //       return doc.getPage(pageNum).then(function (page) {
         //       console.log("# Page " + pageNum);
         //       const viewport = page.getViewport({ scale: 1.0 });
         //       console.log("Size: " + viewport.width + "x" + viewport.height);
         //       console.log();
         //       return page
         //          .getTextContent()
         //          .then(function (content) {
         //             // Content contains lots of information about the text layout and
         //             // styles, but we need only strings at the moment
         //             const strings = content.items.map(function (item) {
         //             return item.str;
         //             });
         //             console.log("## Text Content");
         //             console.log(strings.join(" "));
         //             // Release page resources.
         //             page.cleanup();
         //          })
         //          .then(function () {
         //             console.log();
         //          });
         //       });
         //    };
         //    // Loading of the first page will wait on metadata and subsequent loadings
         //    // will wait on the previous pages.
         //    for (let i = 1; i <= numPages; i++) {
         //       lastPromise = lastPromise.then(loadPage.bind(null, i));
         //    }
         //    return lastPromise;
         // })
         // .then(
         //    function () {
         //       console.log("# End of Document");
         //    },
         //    function (err) {
         //       console.error("Error: " + err);
         //    }
         // );
         //2-pdf-lib 加载没有标记模板pdf(实际使用应该创建一个新的空白pdf),画入 替换好的text
         
         const {PDFDocument} = require("pdf-lib")
         const fs = require("fs")
         const pdfBuffer=fs.readFileSync('D:/Ricardo/documentationProject/my-project/doc/Reporaaaaaat1234.pdf');
         const pdfDoc = await PDFDocument.load(pdfBuffer)
         const form =pdfDoc.getForm()
         console.log("form:",form)
         // console.log('pdfDoc:',pdfDoc)
         // let pages=pdfDoc.getPages()
         // for(let page of pages){
         //    // console.log("pdf page:",page)
         // }
         ctx.body={status:1,data:"success"}
      }catch(err){
         console.log("err:",err)
         ctx.body={status:0,msg:JSON.stringify(err)}
      }
   }
   static mkdirsSync (dirpath) { //同步创建目录
       try
       {
           if (!fs.existsSync(path.dirname(dirpath))) {
               fileUtil.mkdirsSync(path.dirname(dirpath));
            }
            fs.mkdirSync(dirpath);
       }catch(e)
       {
            console.log("create director fail! path=" + dirpath +" errorMsg:" + e);        
       }
   }
   
   static async compressImg ({url,imgSavePath,quality=80,imgWidth=100,imgHeight=100}){
      /**
       * url -原圖片url
       * imgSavePath -壓縮圖片另存文件路徑，如c:/test
       * quality-壓縮質量百分比 不傳默認80
       * imgWidth-壓縮圖片寬度 不傳默認100
       * imgHeight -壓縮圖片高度 不傳默認100
       */
       try{
           if(!fs.existsSync(url)){
               return false
           }
           
         //   await this.mkdirsSync(imgSavePath)
           let reduceFn=()=>{
              if(!fs.existsSync(imgSavePath)){
                 fs.mkdirSync(imgSavePath)
                 return reduceFn()
              }
           }
           await reduceFn()

           const ImageName=url.split('/')[url.split('/').length-1];
           const ImageNameAry=ImageName.split('.')
           let isNeedCompress=!fs.existsSync(imgSavePath+'/'+ImageNameAry[0]+'_s.'+ImageNameAry[1])
           if(isNeedCompress){
               await images(url).resize(imgWidth,imgHeight).save(imgSavePath+'/'+ImageNameAry[0]+'_s.'+ImageNameAry[1],{quality:quality})
           }    
           return true
       }catch(err){
           console.log('compressImg err:',err)
           return err
       } 
   }

   static async testImage6(ctx){
      await images('D:/Henderson/testChqDomcument/pdfImage/0a35de30-e76f-11ea-a4d4-056f68b89ee6.jpg')
            .save('D:/Henderson/testChqDomcument/pdfImage/outFile/0a35de30-e76f-11ea-a4d4-056f68b89ee6.jpg',{operation:50})
      ctx.body="test"
   }

   static async testImage(ctx){
      let readFileList=function(path,filesList){
         var files = fs.readdirSync(path);
         files.forEach(function (itm, index) {
             var stat = fs.statSync(path + itm);
             if (stat.isDirectory()) {
             //递归读取文件
                 readFileList(path + itm + "/", filesList)
             } else {  
                 var obj = {};//定义一个对象存放文件的路径和名字
                 obj.path = path;//路径
                 obj.filename = itm//名字
                 filesList.push(obj);
             }
         })
      }
      let getFiles={
         //获取文件夹下的所有文件
         getFileList: function (path) {
            var filesList = [];
            readFileList(path, filesList);
            return filesList;
        },
        //获取文件夹下的所有图片
        getImageFiles: function (path) {
            var imageList = [{
               "groupName":"私人電梯大堂/燈槽/明角/不平滑 測試測試測試測試測試測試測試測試測試測試測試測試測試測試",
               "Y":10,
               "groupNameStyle":{"width":45,"fontSize":12},
               "imgsData":{
                  "width":100,
                  "height":100,
                  "itemX":36,
                  quality:80,
                  saveImgPath:'D:/Henderson/testChqDomcument/pdfImage/outFile/test01/test02/test03',
                  "imgAry": []
               }
            }];
            // let imageList=[]
            let dataIndex=0;
            this.getFileList(path).forEach((item,index) => {
                var ms = image(fs.readFileSync(item.path + item.filename));
               //  ms.mimeType && (imageList.push({
               //     path:item.path,
               //     filename:item.filename
               //  }))
                ms.mimeType && (imageList[dataIndex].imgsData.imgAry.push({
                   url:item.path + item.filename,
                   link:'www.baidu.com',
                   title:''
                }))
                if(ms.mimeType && index !== 0 && index%4 === 0){
                  imageList.push({
                     "groupName":"私人電梯大堂/燈槽/明角/不平滑 測試測試測試測試測試測試測試測試測試測試測試測試測試測試",
                     "Y":10,
                     "groupNameStyle":{"fontSize":12},
                     "imgsData":{
                        "width":100,
                        "height":100,
                        "itemX":36,
                        quality:80,
                        saveImgPath:'D:/Henderson/testChqDomcument/pdfImage/outFile/test01/test02/test03',
                        "imgAry": []
                     }
                  })
                  dataIndex++;
                }
            });
            return imageList;
         }
      }
      
      let files=getFiles.getImageFiles('D:/Henderson/testChqDomcument/pdfImage/')
      let buffs=[]

      // for(let i =0;i<files.length;i++){
      //    let buff= await images(files[i].path+files[i].filename)
      //      .size(150,150)
      //    //   .save(files[i].path+'outFile/'+files[i].filename,{operation:80})
      //      .encode('jpg',{quality:80})
      //    buffs.push(buff)
      // }
      // console.log("buffs:",buffs)
      // await fileController.compressImg({
      //    url:'D:/Henderson/testChqDomcument/pdfImage/0a35de30-e76f-11ea-a4d4-056f68b89ee6.jpg',
      //    imgSavePath:'D:/Henderson/testFile'
      // })
      files=files.concat(files,files)

      ctx.body={data:files}
   }
   static async testImage2(ctx){
      // const imagemin=require('imagemin')
      // // const imageminJpegtran = require('imagemin-jpegtran');
      // // const imageminMozjpeg = require('imagemin-mozjpeg'); 
      // const imageminJpegRecompress = require('imagemin-jpeg-recompress');
      // let buff= await fs.readFileSync('D:/jacky/documentationProject/my-project/assets/img/j01.jpg')
      // const files=await imagemin(
      //    ['D:/jacky/documentationProject/my-project/assets/img/j01.jpg'],
      //    {
      //       destination:'/my-project/assets/img/test',
      //       plugins:[imageminJpegRecompress({
      //          quality:50
      //       })]
      //    }
      // )
      // {
      //    plugins:[imageminMozjpeg({
      //       quality:30
      //    })]
      // }
      // const files=await imagemin(
      //    ['D:/jacky/documentationProject/my-project/assets/img/j01.jpg'],
      //    {
      //       destination:'D:/jacky/documentationProject/my-project/assets/img/test'
      //    }
      // )
      console.log(buff)
      console.log(files)
      ctx.body={}
      // fs.writeFileSync('D:/jacky/documentationProject/my-project/assets/img/test/test.jpg',files[0].data)
   }
   static async testImage3(ctx){
      const {compress}=require('compress-images/promise')
      const result= await compress({
         source:'D:/jacky/documentationProject/my-project/assets/img/j01.jpg',
         enginesSetup:{
            jpg: { engine: 'mozjpeg', command: ['-quality', '60']},
         }
      })
      console.log("result:",result)

      ctx.body='test'
   }

   static async testImage4(ctx){
      const tinify=require('tinify');
      tinify.key="test"
      let source=tinify.fromFile("D:/jacky/documentationProject/my-project/assets/img/j01.jpg");
      let resized=source.resize({
         method:'fit',
         width:80,
         height:80
      });
      resized.toFile('D:/jacky/documentationProject/my-project/assets/img/test/test.jpg');
      ctx.body="test"
   }
   
}
module.exports = fileController;

// //test
// fileController.testImage();



