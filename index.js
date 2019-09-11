var xlsx = require('node-xlsx');
var fs = require("fs");

//配置区域 开始
var filePrefix = 'abc';//前缀
var readFile = 'Record.xlsx';//读取的文件名
var fieldmap = new Map();//filed的map，key为label，value为type，默认认为apiname为label去掉空格变为下划线，并截取40位长度 +'__c'
fieldmap.set('Group Name','string');
fieldmap.set('External Id','string');
fieldmap.set('Object Name','string');
fieldmap.set('Territory Type','string');
fieldmap.set('Record Type','string');
// 配置区域 结束


var excelData = xlsx.parse("./inputFile/" + readFile);
var sheet1 = excelData[0].data;


if(sheet1.length > 1){
    var header = sheet1[0];
    console.log('header:',header);
    for(var i = 1; i < sheet1.length; i++ ){
        var data = sheet1[i];
        console.log('data:',data);
        write(data,header);
    }
    // writeMdFile();
}


function write(data,header){
    var fileName = filePrefix + '.' + data[header.indexOf('Group Name')].replace(/ /g, '_') + '.mdt';
    console.log("准备写入文件 :" + fileName);
    fs.open('./outputFile/' + fileName, 'w', function(err, fd) {
        if (err) {
            return console.error(err);
        }
        console.log("文件打开成功！");    
        
        var content = '<?xml version="1.0" encoding="UTF-8"?>\n';
        content += '<CustomMetadata xmlns="http://soap.sforce.com/2006/04/metadata" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">\n';
        content += '\t<label>' + data[header.indexOf('Group Name')] + '</label>\n';
        content += '\t<protected>false</protected>\n';
    
        content += '\t<values>\n';
        content += '\t\t<field>AccessLevel__c</field>\n';
        content += '\t\t<value xsi:type="xsd:string">Read</value>\n';
        content += '\t</values>\n';
    
        for(var fieldLabel of fieldmap.keys()){
            var fieldApi = fieldLabel.length > 40?fieldLabel.substring(0,40) :fieldLabel;
            fieldApi = fieldApi.replace(/ /g, '_') ;
            fieldApi += '__c';

            var type =  fieldmap.get(fieldLabel).toLowerCase().trim();

            content += '\t<values>\n';
            content += '\t\t<field>' + fieldApi + '</field>\n';
            content += '\t\t<value xsi:type="xsd:' + type + '">' + data[header.indexOf(fieldLabel)] + '</value>\n';
            content += '\t</values>\n';
        }
    
        content += '</CustomMetadata>';
        console.log('content:',content);
        var buffer = new Buffer(content);

        fs.write(fd, buffer, 0, buffer.length, 0, function (err, written, buffer) {
            console.log('written:',written.toString());
            fs.close(fd, function(err){
                if (err){
                    console.log(err);
                } 
                console.log("文件关闭成功");
            });
        });
     });
}
