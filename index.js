var fs = require('fs');
const path = require('path');
var fse = require('fs-extra');
var unzip = require('unzip');
var xml2js = require('xml2js');
var jsonfile = require('jsonfile');
var pointer = require('json-pointer');
var parser = new xml2js.Parser()

exports.extractComments = function(filepath) {    
      

	if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/tmp/'+filename, function(err){
                    fs.rename(__dirname+'/tmp/'+filename, __dirname+'/tmp/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/tmp/'+newFile).pipe(unzip.Extract({ path: __dirname+'/tmp/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/tmp/'+filename+'/word/comments.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to have any comments");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = 'temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                                var comments = [];

                                                for(i=0; i<2000; i++){

                                                try{
                                                    var test = pointer.get(jsonData, '/w:comments/w:comment/'+i+'/w:p/0/w:r/1/w:t');
                                                    var newComment = test[0].replace(/\W/g, ' ');
                                                    comments[i] = newComment; 
                                                }catch(e){
                                                    return comments; //no more comments to parse, return 
                                                    
                                                 }
                                                }
                                            });
                                        });

                                    });
                                }
                            });
                        });
                    });
                });
          } else {
          return console.log("The file you are passing into the function is not a 'docx' file");
      }



}