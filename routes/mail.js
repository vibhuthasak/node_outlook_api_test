var express = require('express');
var router = express.Router();
var authHelper = require('../helpers/auth');
var graph = require('@microsoft/microsoft-graph-client');
var fs = require('fs');
const parse5 = require('parse5');
// const ExportToCsv = require('export-to-csv');
// import { ExportToCsv } from 'export-to-csv';

var logger = fs.createWriteStream('log.txt', {
    flags: 'a' // 'a' means appending (old data will be preserved)
})

router.get('/', async function(req, res, next) {
    let parms = { title: 'Inbox', active: { inbox: true } };

    const accessToken = await authHelper.getAccessToken(req.cookies, res);
    const userName = req.cookies.graph_user_name;
    var messagesForCsv = [];

    if (accessToken && userName) {
        parms.user = userName;

        // Initialize Graph client
        const client = graph.Client.init({
            authProvider: (done) => {
                done(null, accessToken);
            }
        });

        try {
            // Get the 10 newest messages from inbox
            const result = await client
                .api('/me/mailfolders/inbox/messages')
                .top(1000)
                .select('body')
                .orderby('receivedDateTime DESC')
                .get();

            parms.messages = result.value;
            // console.log(parms.messages.length);
            for (var j of parms.messages){
                // console.log(i.body.content)
                let content = j.body.content;
                // let doc = parse5.parse(content);
                let cleanText = content.replace(/<\/?[^>]+(>|$)/g, "");
                // logger.write(cleanText.trim() + '\n');
                cleanText = cleanText.trim();
                let lines = cleanText.split('\n');
                // logger.write('\n');
                var re = /Subject/;
                var subjectList = 0;
                var confidential = `This e-mail is confidential`;
                var confidentialHit = false;
                var writeContent = '';
                var subjectContent = '';
                var messagesForCsvInd = {};
                // Find Subjects
                for (var i = 0 ; i < lines.length ; i ++) {
                    if(lines[i].match(re)) {
                        console.log(lines[i])
                        subjectList += 1;
                    }
                }
                console.log(subjectList);
                console.log('__________________________________________')

                for(var i = 0;i < lines.length;i++) {
                    // console.log(lineSplit[0])
                    if(lines[i].match(re)) {
                        console.log(lines[i])
                        subjectList -= 1;
                    }
                    
                    if (lines[i].match(re) && subjectList === 0) {
                        subjectContent = lines[i].replace(/&(nbsp|amp|quot|lt|gt);/g, '')
                    }

                    // console.log(subjectList)
                    if (lines[i].includes(confidential)) {
                        confidentialHit = true;
                    }  
                    let lineSplit = lines[i].split(' ');
                    if (subjectList === 0) {
                        if ((lines[i].length > 1) && (lineSplit[0] != '&nbsp;\r') && (lineSplit[0] != 'From:') && (lineSplit[0] != 'Date:') && (lineSplit[0] != 'From:') && (lineSplit[0] != 'Sent:') && (lineSplit[0] != 'Cc:') && (lineSplit[0] != 'To:')) {
                            if (confidentialHit === false) {
                                writeContent += (lines[i].replace(/&(nbsp|amp|quot|lt|gt);/g, '') + '\n');
                            }
                        }
                    }
                }
                // logger.write("SUBJECTBY " + subjectContent)
                // logger.write(writeContent);
                messagesForCsvInd['Subject'] = subjectContent.replace(/Subject:\s*/g, '');
                writeContent = writeContent.replace(subjectContent, '');
                messagesForCsvInd['Message'] = writeContent.replace(/\r|\n/g, '')
                messagesForCsv.push(messagesForCsvInd);
                // logger.write('\n--------------------------------------------------------------------------------\n');
            }
            // const csvExporter = new ExportToCsv(options);
 
            // csvExporter.generateCsv(messagesForCsvInd);
            console.log(messagesForCsv);
            var json = JSON.stringify(messagesForCsv);
            fs.writeFile('emails.json', json, 'utf8', function(){
                console.log('Function Added');
            });
            res.render('mail', parms);
        } catch (err) {
            parms.message = 'Error retrieving messages';
            parms.error = { status: `${err.code}: ${err.message}` };
            parms.debug = JSON.stringify(err.body, null, 2);
            res.render('error', parms);
        }

    } else {
        // Redirect to home
        res.redirect('/');
    }
});

module.exports = router;