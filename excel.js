// const fs = require("fs");
const path = require('path');
const xlsx = require('node-xlsx')
const koaBody = require('koa-body')
const static = require('koa-static');

const Koa = require('koa');
const router = require('koa-router')()

const app = new Koa();

//设置根目录
// var root = './data';
var dataArr = [];

app.use(koaBody({
    multipart: true,
    formidable: {
        maxFileSize: 200*1024*1024    // 设置上传文件大小最大限制，默认2M
    }
}));
app.use(router.routes());

router.post('/upload', async (ctx, next) => {
    await next();
    // const files = ctx.request.body
    const files = ctx.request.files.file

    for (let index = 0; index < files.length; index++) {
        const file = files[index];
        parseExcel(file.path, file.name);
    }
    const buff = writeExcel();
    ctx.length = Buffer.byteLength(buff);

    ctx.set("Content-Disposition", "attachment; filename=" + "o2olog.xlsx")
            
    dataArr = [];
    ctx.body = buff;
    
})

app.use(static( path.join(__dirname, 'www') ));

app.listen(3000);



function parseExcel(filepath,filename) {
    //读取文件内容
    const obj = xlsx.parse(filepath);
    const excelObj=obj[0].data;
    //处理数据
    for(let i=0;i< excelObj.length; i++){
        let row = excelObj[i];
        row.unshift(filename)
        console.log(row)
        dataArr.push(row)
    }
}

function writeExcel(data) {
    const buffer = xlsx.build([
        {
            name:'sheet1',
            data:dataArr
        }
    ]);
    // return buffer
    // var d = new Date().getTime();
    // fs.writeFileSync('./data/'+d+'.xlsx',buffer,{'flag':'w'});
    return buffer
}