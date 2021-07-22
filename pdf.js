pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@2.8.335/build/pdf.worker.js';

const file = document.getElementById("file");
const UploadButton = document.getElementById("upload");
const canvas = document.getElementById('the-canvas');
const ErrorSpan = document.getElementById('error');
const AppBody = document.getElementById('app-body');
const Loader = document.getElementById('loader_parent');
const fileReader = new FileReader();

fileReader.onload = function (e) {
    readPDFFile(new Uint8Array(e.target.result));
};

async function readPDFFile(pdf_data) {

    AppBody.style.display = "none";
    Loader.style.display = "flex";

    const pdf = await pdfjsLib.getDocument({ data: pdf_data }).promise;

    const page_count = pdf.numPages;

    // Run a batch operation against the Word object model.
    await Word.run(async function (context) {

        const body = context.document.body;

        const Pages = document.getElementById('Pages').textContent;

        let Page = [];
        if(Pages) {
            if(/^\d+-\d+$/.test(Pages)) {
                let [StartPage, EndPage] = Pages.split("-");
                StartPage = parseInt(StartPage);
                EndPage = parseInt(EndPage);
                if(StartPage < EndPage && EndPage <= page_count) {
                    for(let i = StartPage; i <= EndPage; i++) {
                        Page.push(i);
                    }
                }
            } else if(/^\d+(,\d+)+$/.test(Pages)) {
                const PagesList = Pages.split(",");
                for(let i = 0; i < PagesList.length; i++) {
                    Page.push(parseInt(PagesList[i]));
                }
            } else {
                Page.push(parseInt(Pages));
            }
            
            Page = Page.filter(item => item && item <= page_count)

            if(Page.length === 0) {
                Page.push(1);
            }

            Page.forEach(function(item) {
                const page = await pdf.getPage(item);
                const viewport = page.getViewport({ scale: document.getElementById('Scale').textContent || 1.5 });
        
        
                const canvasContext = canvas.getContext('2d');
                canvas.height = viewport.height;
                canvas.width = viewport.width;
        
        
                await page.render({ canvasContext, viewport }).promise;
                body.insertInlinePictureFromBase64(canvas.toDataURL('image/jpeg').replace(/^data:image\/\w+;base64,/, ""), "End");
                body.insertBreak("Page", "End");
                await context.sync();
            })
            
        } else {
            for (let i = 1; i <= page_count; i++) {

                const page = await pdf.getPage(i);
                const viewport = page.getViewport({ scale: document.getElementById('Scale').textContent || 1.5 });
        
        
                const canvasContext = canvas.getContext('2d');
                canvas.height = viewport.height;
                canvas.width = viewport.width;
        
        
                await page.render({ canvasContext, viewport }).promise;
                body.insertInlinePictureFromBase64(canvas.toDataURL('image/jpeg').replace(/^data:image\/\w+;base64,/, ""), "End");
                body.insertBreak("Page", "End");
                await context.sync();
    
            }
        }


        AppBody.style.display = "flex";
        Loader.style.display = "none";


        console.log('Added base64 encoded text to the end of the document body.');

    }).catch(function (error) {
        console.log(error);
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        }
    });

    file.value = "";
}

file.onchange = function (e) {

    ErrorSpan.textContent = "";

    let PDFFile = file.files[0];
    if (!PDFFile) return;

    if (PDFFile.type !== "application/pdf") {
        file.value = "";
        ErrorSpan.textContent = "Please select a PDF file.";
        return;
    }

};

console.log(1);

UploadButton.onclick = function () {

    console.log("Clicked");

    ErrorSpan.textContent = "";

    let PDFFile = file.files[0];
    if (!PDFFile) {
        ErrorSpan.textContent = "Please select a PDF file.";
    };

    fileReader.readAsArrayBuffer(PDFFile);
};


console.log(2);