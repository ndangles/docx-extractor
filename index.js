var fs = require('fs');
const path = require('path');
var fse = require('fs-extra');
var unzip = require('unzip');
var xml2js = require('xml2js');
var jsonfile = require('jsonfile');
var pointer = require('json-pointer');
var parser = new xml2js.Parser()





exports.templateUsed = function(filepath, callback){

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/tmp/'+filename, function(err){
                    fs.rename(__dirname+'/tmp/'+filename, __dirname+'/tmp/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/tmp/'+newFile).pipe(unzip.Extract({ path: __dirname+'/tmp/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/tmp/'+filename+'/docProps/app.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to specify the template used in its xml");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/tmp/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                                

                                                try{
                                                    var templateUsed = pointer.get(jsonData, '/Properties/Template');
                                                    fse.emptyDir(__dirname+'/tmp/');
                                                    return callback(templateUsed); 
                                                }catch(e){
                                                    
                                                    return console.log("Error occured trying to get the template used. If you are seeing this error and cannot resolve the issue, contact me at nicholasdangles@gmail.com");
                                                 }
                                                
                                            });
                                        });

                                    });
                                }
                            });
                        });
                    });
                    fse.emptyDir(__dirname+'/tmp/');
                });
                
          } else {
          return console.log("Can't get the template used. The file you are passing into the function is not a 'docx' file");
      }
}




exports.numberPages = function(filepath, callback){

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/tmp/'+filename, function(err){
                    fs.rename(__dirname+'/tmp/'+filename, __dirname+'/tmp/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/tmp/'+newFile).pipe(unzip.Extract({ path: __dirname+'/tmp/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/tmp/'+filename+'/docProps/app.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to have a last modified time in its xml");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/tmp/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                                

                                                

                                                try{
                                                    var numberPages = pointer.get(jsonData, '/Properties/Pages');
                                                    return callback(numberPages);
                                                }catch(e){
                                                    
                                                    return console.log("Error occured trying to get last time modified. If you are seeing this error and cannot resolve the issue, contact me at nicholasdangles@gmail.com");
                                                 }
                                                
                                            });
                                        });

                                    });
                                }
                            });
                        });
                    });
                    fse.remove(__dirname+'/tmp/');
                });
                
          } else {
          return console.log("Can't get last time modified. The file you are passing into the function is not a 'docx' file");
      }
}



exports.lastModified = function(filepath, callback){

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/lm/'+filename, function(err){
                    fs.rename(__dirname+'/lm/'+filename, __dirname+'/lm/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/lm/'+newFile).pipe(unzip.Extract({ path: __dirname+'/lm/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/lm/'+filename+'/docProps/core.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to a last modified time in its xml");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/lm/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                          

                                                try{
                                                    var lastModified = pointer.get(jsonData, '/cp:coreProperties/dcterms:modified/0/_');
                                                    return callback(lastModified);
                                                }catch(e){
                                                    
                                                    return console.log("Error occured trying to get last time modified. If you are seeing this error and cannot resolve the issue, contact me at nicholasdangles@gmail.com");
                                                 }
                                                
                                            });
                                        });

                                    });
                                }
                            });
                        });
                    });
                    fse.emptyDir(__dirname+'/lm/');
                });
                
          } else {
          return console.log("Can't get last time modified. The file you are passing into the function is not a 'docx' file");
      }
}

exports.timeCreated = function(filepath, callback){

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/tc/'+filename, function(err){
                    fs.rename(__dirname+'/tc/'+filename, __dirname+'/tc/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/tc/'+newFile).pipe(unzip.Extract({ path: __dirname+'/tc/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/tc/'+filename+'/docProps/core.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to a created time in its xml");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/tc/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                                

                                                try{
                                                    var timeCreated = pointer.get(jsonData, '/cp:coreProperties/dcterms:created/0/_');
                                                    return callback(timeCreated); 
                                                }catch(e){
                                                    
                                                    return console.log("Error occured trying to get time created. If you are seeing this error and cannot resolve the issue, contact me at nicholasdangles@gmail.com");
                                                 }
                                                
                                            });
                                        });

                                    });
                                }
                            });
                        });
                    });
                    fse.emptyDir(__dirname+'/tc/');
                });
                
          } else {
          return console.log("Can't get time created. The file you are passing into the function is not a 'docx' file");
      }
}




exports.getRevisionNumber = function(filepath, callback){

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/tmp/'+filename, function(err){
                    fs.rename(__dirname+'/tmp/'+filename, __dirname+'/tmp/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/tmp/'+newFile).pipe(unzip.Extract({ path: __dirname+'/tmp/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/tmp/'+filename+'/docProps/core.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to a revision number");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/tmp/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                          

                                                try{
                                                    var revisionNum = pointer.get(jsonData, '/cp:coreProperties/cp:revision');
                                                    return callback(revisionNum) 
                                                }catch(e){
                                                    
                                                    return console.log("Error occured trying to get revisionNum. If you are seeing this error and cannot resolve the issue, contact me at nicholasdangles@gmail.com");
                                                 }
                                                
                                            });
                                        });

                                    });
                                }
                            });
                        });
                    });
                    fse.remove(__dirname+'/tmp/');
                });
                
          } else {
          return console.log("Can't get revision number. The file you are passing into the function is not a 'docx' file");
      }
}






exports.lastModifiedBy = function(filepath, callback){

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/tmp/'+filename, function(err){
                    fs.rename(__dirname+'/tmp/'+filename, __dirname+'/tmp/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/tmp/'+newFile).pipe(unzip.Extract({ path: __dirname+'/tmp/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/tmp/'+filename+'/docProps/core.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to a specified author");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/tmp/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                        

                                                try{
                                                    var author = pointer.get(jsonData, '/cp:coreProperties/cp:lastModifiedBy');
                                                    return callback(author) 
                                                }catch(e){
                                                    
                                                    return console.log("Error occured trying to get author name. If you are seeing this error and cannot resolve the issue, contact me at nicholasdangles@gmail.com");
                                                 }
                                                
                                            });
                                        });

                                    });
                                }
                            });
                        });
                    });
                    fse.remove(__dirname+'/tmp/');
                });
                
          } else {
          return console.log("Can't get author name. The file you are passing into the function is not a 'docx' file");
      }
}



exports.getAuthor = function(filepath, callback){

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/tmp/'+filename, function(err){
                    fs.rename(__dirname+'/tmp/'+filename, __dirname+'/tmp/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/tmp/'+newFile).pipe(unzip.Extract({ path: __dirname+'/tmp/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/tmp/'+filename+'/docProps/core.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to have a specified author");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/tmp/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                                


                                                try{
                                                    var author = pointer.get(jsonData, '/cp:coreProperties/dc:creator');
                                                    return callback(author) 
                                                }catch(e){
                                                    
                                                    return console.log("Error occured trying to get author name. If you are seeing this error and cannot resolve the issue, contact me at nicholasdangles@gmail.com");
                                                 }
                                                
                                            });
                                        });

                                    });
                                }
                            });
                        });
                    });
                    fse.remove(__dirname+'/tmp/');
                });
                
          } else {
          return console.log("Can't get author name. The file you are passing into the function is not a 'docx' file");
      }
}




exports.extractComments = function(filepath, callback) {    
      
    var comments = [];

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
                                        var file = __dirname+'/tmp/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                                

                                                for(i=0; i<2000; i++){

                                                try{
                                                    var test = pointer.get(jsonData, '/w:comments/w:comment/'+i+'/w:p/0/w:r/1/w:t');
                                                    var newComment = test[0].replace(/\W/g, ' ');
                                                    comments[i] = newComment; 
                                                }catch(e){
                                                    
                                                    return callback(comments);
                                                 }
                                                }
                                            });
                                        });

                                    });
                                }
                            });
                        });
                    });
                    fse.remove(__dirname+'/tmp/');
                });
                
          } else {
          return console.log("Can't get comments. The file you are passing into the function is not a 'docx' file");
      }

}



exports.getHyperlinks = function(filepath, callback) {    
      
    var hyperlinks = [];

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/tmp/'+filename, function(err){
                    fs.rename(__dirname+'/tmp/'+filename, __dirname+'/tmp/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/tmp/'+newFile).pipe(unzip.Extract({ path: __dirname+'/tmp/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/tmp/'+filename+'/word/document.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to have any hyperlinks");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/tmp/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                                

                                                for(i=0; i<100; i++){

                                                try{
                                                    var test = pointer.get(jsonData, '/w:document/w:body/w:p/'+i+'/w:hyperlink/w:r/w:t');
                                                    if(test!=null){
                                                        hyperlinks[i] = test;
                                                    } 
                                                }catch(e){
                                                    
                                                    return callback(hyperlinks);
                                                 }
                                                }
                                            });
                                        });

                                    });
                                }
                            });
                        });
                    });
                    fse.remove(__dirname+'/tmp/');
                });
                
          } else {
          return console.log("Can't get hyperlinks. The file you are passing into the function is not a 'docx' file");
      }

}