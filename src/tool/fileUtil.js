
import fs from 'fs'
import path from 'path';
class fileUtil{
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
}

module.exports = fileUtil;