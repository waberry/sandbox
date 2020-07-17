import fs from 'fs';
import pkg from 'docx';
import sesTransport from 'nodemailer-ses-transport';
import mailer from 'nodemailer';
//import promise from 'promise';

const { Document, Paragraph, TextRun, Packer } = pkg;
const {writeFileSync} = fs;


// Documents contain sections, you can have multiple sections per document, go here to learn more about sections
// This simple example will only contain one section


function docxcreator(email=null, rawtxt="testtextlore ipsum..blabla") {
   
    // Create document
    const doc = new pkg.Document({
        creator: "Dolan Miu",
        description: "My extremely interesting document",
        title: "My Document",
    });

    doc.addSection({
        properties: {},
        children: [
            new Paragraph({
                children: [
                    new TextRun(rawtxt)
                ],
            }),
        ],
    });
    
    // Used to export the file into a .docx file
    Packer.toBuffer(doc).then((buffer) => {
        writeFileSync('testing.docx', buffer);
    });
}

function   docxsender()
{
    let transporter = mailer.createTransport(
		sesTransport({
			accessKeyId: 'AKIAIOFXDAKWV4JFULKQ',
			secretAccessKey: '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',
			rateLimit: 5,
			region: 'eu-west-1'
		})
    );
    
    // setting mail options
    let message = {
		from: 'Automatly <team@automatly.co>',
		to: 'waberryd@gmail.com',
		subject: 'docxmailer testing',
        //html: htmlcontent,
        attachements: [
            { 
            // stream as an attachment
            filename: 'testing.docx',
            content: fs.createReadStream('testing.docx')
            },
        ],
    };
    transporter.sendMail([message, console.log('end of program!\n')]);
}
// var prom = docxcreator();
// prom.then(docxsender, console.log("----------Error occured during orchestration-------"));

var promise = new Promise(function(resolve, reject) {
    docxcreator();
    // Open a log file
    
    if(fs.existsSync('./testing.docx')){
      docxsender();
    }
    else {
      reject(Error("It broke"));
    }
  });
