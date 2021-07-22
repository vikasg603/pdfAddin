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

        for (let i = 1; i <= page_count; i++) {

            const page = await pdf.getPage(i);
            const viewport = page.getViewport({ scale: 1.5 });
    
    
            const canvasContext = canvas.getContext('2d');
            canvas.height = viewport.height;
            canvas.width = viewport.width;
    
    
            await page.render({ canvasContext, viewport }).promise;
            body.insertInlinePictureFromBase64(canvas.toDataURL('image/jpeg').replace(/^data:image\/\w+;base64,/, ""), "End");
            body.insertBreak("sectionNext", "End");
            await context.sync();

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

UploadButton.onclick = function () {

    ErrorSpan.textContent = "";

    let PDFFile = file.files[0];
    if (!PDFFile) {
        ErrorSpan.textContent = "Please select a PDF file.";
    };

    fileReader.readAsArrayBuffer(PDFFile);
};
