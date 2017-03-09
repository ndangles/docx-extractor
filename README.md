Node.js: docx-extractor
=================

This module allows you to extract comments in 'docx' files and other data stored such as author name, hyperlinks, modification times, etc.


Why?
----

It was a requirement of a job I was working on and I could not find any modules to do this. It was a pain to have to try and create it at the time of the job so I figured I would just make a module incase anyone else ever needed it.

Please feel free to make any comments or suggestions. It can always be improved so I accept any and all feedback.




Installation
------------

npm install --save docx-extractor



Usage
-----

docx-extractor is a module that will let you extract the comments embedded into the xml of 'docx' files as well as other data.

Every method uses the same format as below. Just replace "someMethod" below with the name of the method you want to use.

Example:

```js
var dxe = require('docx-extractor');

dxe.someMethod('myfile.docx', function(data){
    console.log(data)
});
```

 Methods
 -------

 - `extractComments(file, callback)`

 		Extracts the comments embedded into the xml of 'docx' files. It returns all of the comments as an array.


 
 - `getHyperlinks(file, callback)`

 		Returns all hyperlinks in the docx file as an array



 - `getRevisionNumber(file, callback)`

 		Returns the number of times the document was modified and saved.



 - `numberPages(file, callback)`

 		Returns number of pages in the docx file



 - `getAuthor(file, callback)`

 		Returns the original author of the document. This will be the name of the user account on the machine it was created from.



 - `lastModifiedBy(file, callback)`

 		Returns the last person that modified the document. Again, this returns the name of the computer's user account that was in use at the time.



 - `timeCreated(file, callback)`

 		Returns the time the docx file was originally created. Example: 2017-03-08T17:40:00Z



- `lastModified(file, callback)`
	
		Returns the time that the document was last modified. Example: 2017-03-09T14:23:00Z



- `templateUsed(file, callback)`

		Returns the template that was used when creating the document. Example: "Normal.dotm" is the template name of when someone starts a docx file with the blank template.
		


 Future
 ------

 I plan to add more methods in the near future for extracting and parsing other specific data out of the 'docx' xml files. 

 My next plan is to add word, character and paragraph count. Although this information is stored in the xml, microsoft recommends against it because it can be inconsistent with the values that the client application(Word gui) shows.