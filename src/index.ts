
import * as docx from './docx-preview.js';
import * as fs from 'fs';
const jsdom = require("jsdom");
const { JSDOM } = jsdom;

const docxOptions = Object.assign(docx.defaultOptions, {
    debug: true,
    experimental: true,
    //debug
    inWrapper: false, 
    renderComments: true,
    hideWrapperOnPrint: false,
    ignoreLastRenderedPageBreak: true,
});

//const container = document.querySelector("#document-container");
const fileName = '/home/franemar/Library/Software-Engineering/Programming-Languages/Functional-Programming/Imperative_to_functional_programming_succinctly.docx';
enum Format {
    "docx" = "Microsoft Word/LibreOffice Doc",
    "epub" = "Electronic Publication"
}
const outputPath = '/home/franemar/Temp/Imperative_to_functional_programming_succinctly.html'
//const fileInput = document.querySelector("#files");
//const loadButton = document.querySelector("#loadButton");
//const testDocuments = document.querySelector("#testDocuments");

/*async function renderDocx(file) {
    currentDocument = file; 

    if (!currentDocument) 
        return;
    
    //// optional, find and convert all tiff images 
    let docxBlob = preprocessTiff(currentDocument);
    // render document
    let res = await docx.renderAsync(docxBlob, container, null, docxOptions)
    // optional - render thumbnails
    renderThumbnails(container, document.querySelector("#thumbnails-container"));
    console.log(res);
}
*/

function convertDocument({ fileName, from, to, outputPath }: 
    { fileName: string; from: Format; to: Format; outputPath: string; }): void {
    fs.readFile(fileName, async (err, data)=> {
        if (err) {
            console.error('Error reading file:', err);
            return;
        }

        //let docxBlob = preprocessTiff(data);

        let parsedDocument = await docx.parseAsync (
            { data, userOptions: docxOptions }
        )

        /*//debug
        fs.promises.writeFile('/home/franemar/Temp/parsedDocument.json',
            JSON.stringify(parsedDocument),
            {
                flag: 'w',
            }
        )
        return
        */

        //const buf = Buffer.from(parsedDocument, 'base64');
        const dom = new JSDOM(`...`, {includeNodeLocations: true});
        let result = await docx.renderDocument(parsedDocument, dom.window, 
            dom.window.document.body, null, docxOptions);
        console.log(result)

        fs.promises.writeFile(outputPath,
            //JSON.stringify(container),
            dom.serialize(),
            {
                flag: 'w',
            }
        )
    });
}

convertDocument({ fileName, from: Format.docx, to: Format.epub, outputPath })
console.log("Document converted and exported to: '%s'", outputPath)
/*fileInput.addEventListener("change", ev => {
    renderDocx(fileInput.files[0]);
    testDocuments.selectedIndex = 0;
});
loadButton.addEventListener("click", ev => renderDocx(fileInput.files[0]));
*/

//const menu = document.querySelector("#optionsMenu");

/*Object.keys(docxOptions).filter(key => !/className/i.test(key)).forEach(function(key) {
    const listItem = document.createElement("li");
    listItem.innerHTML = `
    <div class="dropdown-item">
        <div class="form-check">
            <label class="form-check-name"><input type="checkbox" class="form-check-input" ${docxOptions[key] ? 'checked' : ''}> ${key}</label>
        </div>
    </div>`;

    const checkInput = listItem.querySelector("input");

    checkInput.addEventListener("click", (e) => {
        docxOptions[key] = checkInput.checked;
        renderDocx(currentDocument);
    });

    menu.appendChild(listItem);
});*/

/*container.addEventListener("dragover", ev => ev.preventDefault());
container.addEventListener("drop", ev => {
    ev.preventDefault();
    renderDocx(ev.dataTransfer.files[0]);
});*/

/*for (let testName of ['text','underlines','text-break','line-spacing','numbering','page-layout','table','table-spans','footnote','header-footer','revision','equation'])
{
    var op = document.createElement("option");
    op.label = testName;
    testDocuments.add(op);
}*/

/*testDocuments.addEventListener("change", async e => {
    const selected = testDocuments.selectedOptions[0].label;

    if (selected) {
        let resp = await fetch(`tests/render-test/${selected}/document.docx`);
        renderDocx(await resp.blob());
        fileInput.value = "";
    }
});

document.querySelector("#saveTestButton").addEventListener("click", async () => {
    const file = await showSaveFilePicker({ types: [ { description: "HTML File", accept: { "text/html": [".html"] } }] });
    const stream = await file.createWritable();
    await stream.write(container.innerHTML);
    await stream.close();
});
*/
