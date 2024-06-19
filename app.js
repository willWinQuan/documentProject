
const Koa =require('koa');
const http =require('http');
const KoaBody = require('koa-body');
const KoaLogger = require('koa-logger');
import cors from 'koa2-cors';
const CONFIG=require('./config');
const router =require('./src/routes');
// const render = require('./lib/render');
const views = require('koa-views');
const path = require('path');

const app = new Koa();

app 
    .use(cors({
        origin:"*",
        exposeHeaders: ['WWW-Authenticate', 'Server-Authorization'],
        maxAge: 5,
        credentials: true,
        allowMethods: ['GET', 'POST', 'DELETE'],
        allowHeaders: ['Content-Type', 'Authorization', 'Accept']
    }))
    .use(views(path.join(__dirname, 'src/views'), {
        extension: "html"
    }))
    .use(KoaBody(
        {
            multipart: true,
            jsonLimit: "500mb",
            formidable: {
                maxFileSize: 1024 * 1024 * 1024,
            }
        }))
    .use(KoaLogger())
    .use(router.routes())
    .use(router.allowedMethods())
    .use(async (ctx, next) => {
        try {
            await next()
        } catch (err) {
            ctx.body = { status: 0, message: err.message };
        }
   })


let listenPort = CONFIG.SYSTEM.PROT;
if(CONFIG.SYSTEM.IISNODE){
    listenPort=process.env.PROT;
}
const server = http.createServer(app.callback()).listen(listenPort);
server.timeout = 3 * 60 * 1000;
console.log(`服务启动了：路径为：127.0.0.1:${listenPort}`);

export default app