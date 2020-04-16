var Excel = require('exceljs');
const fs = require('fs');
var path = require('path'); //解析需要遍历的文件夹
var filePath = path.resolve('../shuoming');


var cidArr = ['cid'];
var smArr = ['说明内容'];


//文件遍历方法
function fileDisplay(filePath) {
    //根据文件路径读取文件，返回文件列表
    fs.readdir(filePath, function (err, files) {
        if (err) {
            console.warn(err)
        } else {
            //遍历读取到的文件列表
            files.forEach(function (filename) { //是里面的路径
                //获取当前文件的绝对路径
                var filedir = path.join(filePath, filename); //filedir是新的路径
                //根据文件路径获取文件信息，返回一个fs.Stats对象
                fs.stat(filedir, function (error, stats) {
                    if (error) {
                        console.warn('获取文件stats失败');
                    } else {
                        var isFile = stats.isFile(); //是文件
                        var isDir = stats.isDirectory(); //是文件夹
                        if (isFile) {
                            // 读取文件内容
                            // console.log("num1", filedir)
                            var index = filedir.lastIndexOf(".");
                            //获取后缀
                            var ext = filedir.substr(index + 1);
                            if (ext == "txt") {
                                // console.log(ext)
                                console.log("num2", filedir)
                                cidArr.push(filedir.substring(filedir.indexOf("cid="),filedir.indexOf("cid")+7));
                                var content = fs.readFileSync(filedir,'utf-8');
                                smArr.push(content);
                                // console.log(smArr)
                                // console.log(cidArr)
                            }
                        }
                        if (isDir) {
                            fileDisplay(filedir); //递归，如果是文件夹，就继续遍历该文件夹下面的文件
                        }
                    }
                })
            });
        }
    });
}


function writeXlsx() {
    //创建工作蒲
    var workbook = new Excel.Workbook();
    //添加工作表
    var worksheet = workbook.addWorksheet('My Sheet');
    // 添加列标题并定义列键和宽度
    // 注意：这些列结构只是工作簿构建方便，
    // 除了列宽之外，它们不会完全持久化。
    worksheet.columns = [{
            header: 'cid',
            key: 'cid',
            width: 10
        },
        {
            header: 'shuoming',
            key: 'shuoming',
            width: 32
        }
    ];
    // 按键，字母和从1开始的列号访问各列
    var cidCol = worksheet.getColumn('cid');
    var smCol = worksheet.getColumn('shuoming');
    // 注意：这将覆盖单元格值C1:C2
    // cidCol.header = ['cid'];//标题
    // smCol.header = ['说明内容'];
    // 添加一列新值
    cidCol.values = cidArr;
    smCol.values = smArr;


    // write to a file
    // var workbook = createAndFillWorkbook();
    workbook.xlsx.writeFile("志愿说明.xlsx");
}

async function main() {
    //调用文件遍历方法
    await fileDisplay(filePath);
    console.log(smArr)
    console.log(cidArr)
    setTimeout(function(){
        writeXlsx()
    },2000)
    // await writeXlsx();
}

main()