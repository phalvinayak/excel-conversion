const fileInput = document.querySelector('#upload');
const schema = {
    'Sl.no.': {
        prop: 'slNo',
        type: Number
    },
    'FOSS-Component': {
        prop: 'fossComponent',
        type: String
    },
    'Version': {
        prop: 'version',
        type: String
    },
    'License Name': {
        prop: 'licenseName',
        type: String
    },
    'Copyright': {
        prop: 'copyright',
        type: String
    },
    'License Text': {
        prop: 'licenseText',
        type: String
    },
}

let outputMapping = {};

fileInput.addEventListener('change', e => {
    $('.loader').show();
    const fileName = fileInput.files[0].name;
    const fileExt = /(xlsx|xls)$/i
    if(fileExt.test(fileName)) {
        readXlsxFile(fileInput.files[0], {schema}).then(({rows, errors}) =>  {
            console.log(rows, errors);
            if(errors.length){
                // alert('Unable to map the schema, please validate schema');
                alert(errors.join('\n'));
            } else {
                rows.forEach(row => {
                    const copyrightArr = row.copyright.split('\n');
                    if(outputMapping[row.licenseName]){
                        outputMapping[row.licenseName].push(...copyrightArr);
                    } else {
                        outputMapping[row.licenseName] = copyrightArr;
                    }
                });
                generateWordDocument(outputMapping);
            }
        });
    } else {
        alert('Please select valid Excel file with xls or xlsx extension');
    }
});


async function generateWordDocument(outputMapping){
    const $template = $($('#word-template').text());
    const $wordOutput = $('<div/>', {'id': 'word-output'});
    Object.keys(outputMapping).forEach(key => {
        let $copyRightBlock = $template.clone();
        console.log(key, outputMapping[key]);
        $copyRightBlock.find('.title').text(key);
        let $copyrightArr = outputMapping[key].map(copyright =>
            $('<div/>', {text: copyright})
        );
        $copyRightBlock.find('.copyright').empty().append($copyrightArr);
        $wordOutput.append($copyRightBlock);
    });

    const htmlString = `<!DOCTYPE html>
        <html lang="en">
        <head>
        <meta charset="UTF-8">
        <title>Document</title>
        <style>
            body{font-family: sans-serif;font-size: 12px;white-space: nowrap;line-height: 14px;margin: 5px 10px;padding: 5px;}
            h3.title{color: #4173AE; margin-bottom: 5px}
            .copy-block{margin-bottom: 12px}
        </style>
        </head>
        <body>
        ${$wordOutput[0].outerHTML}
        </body>
        </html>`
    var converted = htmlDocx.asBlob(htmlString, {orientation: 'portrait'});
    saveAs(converted, "output.docx");
    $('.loader').hide();
}