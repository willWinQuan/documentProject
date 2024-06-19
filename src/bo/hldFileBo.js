import {fileUtil} from '../tool';

class hldFileBO {  
   static async getFile(opt){
      let fileRes;
      switch (opt.fileType){
         case 'docx':
           const hsDocxUtil=require('../tool/hsNewDocxUtil');
           fileRes=new hsDocxUtil(opt).getBuf();
         break;
         case 'excel':
         //   const hsExcelUtil=require('../tool/hsExcelUtil');
         //   fileRes=await new hsExcelUtil(opt.data).creatExcelBufferAsync();
           const hsExcelUtil=require('../tool/hsNewExcelUtil');
           fileRes=await new hsExcelUtil(opt.data).render('buff');
         break;
         case 'pdf':
           const hsPDFUtil=require('../tool/hsPDFUtil');
           opt.data["outFileName"]=opt.outFileName;
           fileRes= await new hsPDFUtil(opt.data).render(opt.dataType); 
         //   console.log("fileRes:",fileRes)
         break;
         case 'docxToPdf':
            const _hsDocxUtil=require('../tool/hsNewDocxUtil');
            const libre=require('libreoffice-convert')
            const { promisify }=require("bluebird")
            const docxRes=new _hsDocxUtil(opt).getBuf(); 
            let lib_convert = promisify(libre.convert)
            let pdfbuf = []
            pdfbuf = await lib_convert(docxRes.data, '.pdf', undefined)
            fileRes={status:1,data:pdfbuf}
         break;
     }
     if(opt.dataType === 'file' && fileRes.status && opt.fileType !== 'pdf'){
         const fs =require('fs');
         if(!fs.existsSync(opt.outFilePath)){
            fileUtil.mkdirsSync(opt.outFilePath)
         }
         // fs.writeFileSync(opt.outFileName,fileRes.data);
         fs.writeFileSync(opt.outFilePath+'/test.pdf',fileRes.data);
         fileRes['data']={
            fileName:opt.outFileName
         }
      }
      return fileRes;
   }

}

module.exports = hldFileBO;