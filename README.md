Node.js: docx-extractor
=================

This module allows you to extract comments in 'docx' files.


Why?
----

It was a requirement of a job I was working on and I could not find any modules to do this. It was a pain to have to try and create it at the time of the job so I figured I would just make a module incase anyone else ever needed it.




Installation
------------

npm install --save docx-extractor



Usage
-----

docx-extractor is a module that will let you extract the comments embedded into the xml of 'docx' files. It returns all of the comments as an array.

Example:

```js
var dxe = require('docx-extractor');

dxe.extractComments('myfile.docx', function(data){
    console.log(data)
});
```

 Future
 ------

 I plan to add more methods in the near future for extracting and parsing other specific data out of the 'docx' xml files.