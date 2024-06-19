

const router =new require('koa-router')();
import controllers from '../controllers';

router.post('/getFile',controllers.hldFile.getFile);
router.post('/testImage',controllers.hldFile.testImage);
router.post('/testImage2',controllers.hldFile.testImage2);
router.post('/testImage6',controllers.hldFile.testImage6);
router.post('/getToken',controllers.hldAuthen.getToken);
router.post('/verifyToken',controllers.hldAuthen.verifyToken);
router.get('/testEditPDF',controllers.hldFile.testEditPDF)
router.get('/demo1',async (ctx)=>{
    await ctx.render('demo1')
})
module.exports=router;

