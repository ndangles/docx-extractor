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
                                                    fse.emptyDir(__dirname+'/tmp/', function(err){
                                                        return callback(templateUsed);
                                                    });
                                                     
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
                });
                
          } else {
          return console.log("Can't get the template used. The file you are passing into the function is not a 'docx' file");
      }
}




exports.numberPages = function(filepath, callback){

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/np/'+filename, function(err){
                    fs.rename(__dirname+'/np/'+filename, __dirname+'/np/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/np/'+newFile).pipe(unzip.Extract({ path: __dirname+'/np/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/np/'+filename+'/docProps/app.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to have a last modified time in its xml");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/np/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                                

                                                

                                                try{
                                                    var numberPages = pointer.get(jsonData, '/Properties/Pages');
                                                    fse.emptyDir(__dirname+'/np/',function(err){
                                                        return callback(numberPages);
                                                    });
                                                    
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
                                                    fse.emptyDir(__dirname+'/lm/', function(err){
                                                        return callback(lastModified);
                                                    });
                                                    
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
                                                    fse.emptyDir(__dirname+'/tc/',function(err){
                                                        return callback(timeCreated); 
                                                    });
                                                    
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
                });
                
          } else {
          return console.log("Can't get time created. The file you are passing into the function is not a 'docx' file");
      }
}




exports.getRevisionNumber = function(filepath, callback){

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/rn/'+filename, function(err){
                    fs.rename(__dirname+'/rn/'+filename, __dirname+'/rn/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/rn/'+newFile).pipe(unzip.Extract({ path: __dirname+'/rn/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/rn/'+filename+'/docProps/core.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to a revision number");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/rn/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                          

                                                try{
                                                    var revisionNum = pointer.get(jsonData, '/cp:coreProperties/cp:revision');
                                                    fse.emptyDir(__dirname+'/rn/', function(err){
                                                        return callback(revisionNum);
                                                    });
                                                    
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
                });
                
          } else {
          return console.log("Can't get revision number. The file you are passing into the function is not a 'docx' file");
      }
}






exports.lastModifiedBy = function(filepath, callback){

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/lmb/'+filename, function(err){
                    fs.rename(__dirname+'/lmb/'+filename, __dirname+'/lmb/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/lmb/'+newFile).pipe(unzip.Extract({ path: __dirname+'/lmb/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/lmb/'+filename+'/docProps/core.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to a specified author");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/lmb/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                        

                                                try{
                                                    var author = pointer.get(jsonData, '/cp:coreProperties/cp:lastModifiedBy');
                                                    fse.emptyDir(__dirname+'/lmb/',function(err){
                                                        return callback(author);
                                                    });
                                                    
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
                });
                
          } else {
          return console.log("Can't get author name. The file you are passing into the function is not a 'docx' file");
      }
}



exports.getAuthor = function(filepath, callback){

    if(filepath.indexOf(".docx")>-1){
                var filename = path.basename(filepath);
                var newFile = filename+'.zip';
                fse.copy(filepath, __dirname+'/ga/'+filename, function(err){
                    fs.rename(__dirname+'/ga/'+filename, __dirname+'/ga/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/ga/'+newFile).pipe(unzip.Extract({ path: __dirname+'/ga/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/ga/'+filename+'/docProps/core.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to have a specified author");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/ga/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                                


                                                try{
                                                    var author = pointer.get(jsonData, '/cp:coreProperties/dc:creator');
                                                    fse.emptyDir(__dirname+'/ga/',function(err){
                                                       return callback(author); 
                                                   });
                                                    
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
                fse.copy(filepath, __dirname+'/ec/'+filename, function(err){
                    fs.rename(__dirname+'/ec/'+filename, __dirname+'/ec/'+newFile, function(err) {
                        fs.createReadStream(__dirname+'/ec/'+newFile).pipe(unzip.Extract({ path: __dirname+'/ec/'+filename})).on('close', function () {
                            fs.readFile(__dirname + '/ec/'+filename+'/word/comments.xml', function(err, data) {
                                 if(err){
                                    
                                    return console.log("This document does not appear to have any comments");
                                
                                } else{
                                    parser.parseString(data, function (err, result) {

                                        parsedData = JSON.stringify(result);
                                        var file = __dirname+'/ec/temp.json';
                                        jsonfile.writeFile(file, parsedData, function(err){
                                            jsonfile.readFile(file, function(err, obj) {
                                                var jsonData = JSON.parse(obj);
                                                

                                                

                                                try{
                                                    for(i=0; i<2000; i++){
                                                        var test = pointer.get(jsonData, '/w:comments/w:comment/'+i+'/w:p/0/w:r/1/w:t');
                                                        var newComment = test[0].replace(/\W/g, ' ');
                                                        comments[i] = newComment;

                                                    } 
                                                }catch(e){
                                                    
                                                    fse.emptyDir(__dirname+'/ec/',function(err){
                                                        return callback(comments);
                                                        
                                                    });
                                                    
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
          return console.log("Can't get comments. The file you are passing into the function is not a 'docx' file");
      }

}



// exports.getHyperlinks = function(filepath, callback) {    
      
//     var hyperlinks = [];

//     if(filepath.indexOf(".docx")>-1){
//                 var filename = path.basename(filepath);
//                 var newFile = filename+'.zip';
//                 fse.copy(filepath, __dirname+'/ghl/'+filename, function(err){
//                     fs.rename(__dirname+'/ghl/'+filename, __dirname+'/ghl/'+newFile, function(err) {
//                         fs.createReadStream(__dirname+'/ghl/'+newFile).pipe(unzip.Extract({ path: __dirname+'/ghl/'+filename})).on('close', function () {
//                             fs.readFile(__dirname + '/ghl/'+filename+'/word/document.xml', function(err, data) {
//                                  if(err){
                                    
//                                     return console.log("This document does not appear to have any hyperlinks");
                                
//                                 } else{
//                                     parser.parseString(data, function (err, result) {

//                                         parsedData = JSON.stringify(result);
//                                         var file = __dirname+'/ghl/temp.json';
//                                         jsonfile.writeFile(file, parsedData, function(err){
//                                             jsonfile.readFile(file, function(err, obj) {
//                                                 var jsonData = JSON.parse(obj);
                                                

//                                                 for(i=0; i<100; i++){

//                                                 try{
//                                                     var test = pointer.get(jsonData, '/w:document/w:body/w:p/'+i+'/w:hyperlink/w:r/w:t');
//                                                     if(test!=null){
//                                                         hyperlinks[i] = test;
//                                                     } 
//                                                 }catch(e){

//                                                     fse.emptyDir(__dirname+'/ghl/',function(err){
//                                                        return callback(hyperlinks);
//                                                         
//                                                    });
                                                    
//                                                  }
//                                                 }
//                                             });
//                                         });

//                                     });
//                                 }
//                             });
//                         });
//                     });
//                 });
                
//           } else {
//           return console.log("Can't get hyperlinks. The file you are passing into the function is not a 'docx' file");
//       }

// }