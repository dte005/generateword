const express = require('express')
const app = express();
const docx = require('docx');
const dotenv = require('dotenv');
//const fs = require('fs');
const { Document, Packer, Paragraph, Table, TableCell, TableRow,TextRun  } = docx;
//import * as docx from "docx";
//import * as fs from "fs";
dotenv.config()

app.get('/', (req, res) => {
  res.send('Hello World!')
});
app.listen(process.env.PORT || 8000);
app.get("/word", (req, res) => {

  console.log("executing....");
  const doc = new Document();

  doc.addSection({
    properties: {},
    children: [
        new Paragraph({
            children: [
                new TextRun("Hello World"),
                new TextRun({
                    text: "Foo Bar",
                    bold: true,
                }),
                new TextRun({
                    text: "\tGithub is the best",
                    bold: true,
                }),
            ],
        }),
    ],
});

  //  const b64string = Packer.toBase64String(doc);
  //  console.log(b64string)
  //  res.setHeader('Content-Disposition', 'attachment; filename=My Document.docx');
  //  res.send(Buffer.from(b64string, 'base64'));
  // Used to export the file into a .docx file
  const b64string = Packer.toBase64String(doc)
  .then(function(result){
    return result;
  });
  //Packer.toBuffer(doc).then((buffer) => {
  //fs.writeFileSync("My Document.docx", buffer);
  //});
  //res.send(Buffer.from(b64string, 'base64'));


  res.setHeader('Content-Disposition', 'attachment; filename=My Document.docx');
  var doc64 = Buffer.from('base64',b64string,);
    res.writeHead(200, {
    'Content-Type': 'application/octet-stream',
    'Content-Length': doc64.length
  });
  res.end(doc64);

  console.log('escreveu base 64');
  //var files = fs.createReadStream("My Document.docx");
  //res.writeHead(200, {'Content-disposition': 'attachment; filename=My Document.docx'}); //here you can add more headers
  //files.pipe(res)
});
