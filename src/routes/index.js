
const router =new require('koa-router')();
// const fileRoute= require('./fileRoute');
import fileRoute from './fileRoute';

router.use('',fileRoute.routes(),fileRoute.allowedMethods());

module.exports = router;