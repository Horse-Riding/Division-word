var fs = require('fs');
var http = require("http");
var officegen = require('officegen');

let AdmZip = require('adm-zip');
let zip = new AdmZip('./demo.zip');
// console.log(zip);
let contentXml = zip.readAsText("word/document.xml");
let str = "";
// console.log(contentXml.match(/<w:t>[\s\S]*?<\/w:t>/ig));

let searchData = contentXml.match(/<w:t[\s\S]*?>[\s\S]*?<\/w:t>/ig); 
// searchData.forEach((item)=>{
//   str += item.slice(5,-6);
// });

// console.log(str);

http.createServer ( function ( request, response ) {
  response.writeHead ( 200, {
    "Content-Type": "application/vnd.openxmlformats-officedocxument.presentationml.presentation",
    'Content-disposition': 'attachment; filename=surprise.pptx'
    });
}).listen ( 3000 );
async function generate(title, obj){
    return new Promise((resolve, reject) => {
        var docx = officegen({
          'type':'docx'
        });
        docx.on('finalize', function(written) {
          console.log(
            'Finish to create a Microsoft Word document.'
          )
        });
        var uptitle = title || '无名';
        obj.forEach((cur) => {
          let pObj = docx.createP();
          pObj.addText(cur);
        })
        var out = fs.createWriteStream ( 'out/' + uptitle + '.docx' ); 
        docx.generate ( out, {
        'finalize': function(data){
            console.log(data);
        },
        'error': reject,
        });

        out.on('finish', function(){
            resolve(true);
        });
    });
}
let docxObj = {};
let currentIndex = '';
searchData.forEach((item)=>{
  let str = '';
  if (item.indexOf('<w:t>') > -1) {
    str = item.slice(5,-6);
  } else {
    // 换页时标签不同单独处理
    str = item.slice(26,-6);
  }
  if (str[0] === '：') {
    docxObj[str.replace(/：/g,"w")] = [];
    currentIndex = str.replace(/：/g,"w");
  } else {
    if (str.indexOf('成果序号') > -1) {
      str.replace(/'成果序号'/g,"");
    } else {
      if (docxObj[currentIndex]) {
        docxObj[currentIndex].push(str);
      }
    }
  }
});
// console.log(docxObj);
// 调用
for (let i in docxObj) {
  generate(i,docxObj[i]);
}
