var fs = require('fs');
var fse = require('fs-extra');
var unzip = require('unzip');
var xml2js = require('xml2js');
var jsonfile = require('jsonfile');
var pointer = require('json-pointer');
var parser = new xml2js.Parser()

exports.extractComments = function(filename) {

	if(filename.indexOf(".docx")>-1){
                var newFile = filename+'.zip';
                fse.copy(__dirname+'tmp/'+filename, __dirname+'tmp/extdata/'+filename, function(err){
                    fs.rename(__dirname+'tmp/extdata/'+filename, __dirname+'tmp/extdata/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'tmp/extdata/'+newFile).pipe(unzip.Extract({ path: __dirname+'tmp/extdata/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/tmp/extdata/'+filename+'/word/comments.xml', function(err, data) {
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
                                                    console.log(comments[i]); 
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