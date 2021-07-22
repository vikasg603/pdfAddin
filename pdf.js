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

const InsertImages = (base64Images) => {
    // Run a batch operation against the Word object model.
    Word.run(async function (context) {

        const body = context.document.body;

        base64Images.forEach(function (base64Image) {
            base64Image = base64Image.replace(/^data:image\/\w+;base64,/, "");
            body.insertInlinePictureFromBase64(base64Image, "End");
        });

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        await context.sync();
        delete AppBody.style.display;
        Loader.style.display = "none";
        console.log('Added base64 encoded text to the beginning of the document body.');
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}

async function readPDFFile(pdf_data) {

    AppBody.style.display = "none";
    delete Loader.style.display;

    const pdf = await pdfjsLib.getDocument({ data: pdf_data }).promise;

    const page_count = pdf.numPages;
    const Images = [];
    for (let i = 1; i <= page_count; i++) {

        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: 1.5 });


        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;


        await page.render({ canvasContext: context, viewport: viewport }).promise
        Images.push(canvas.toDataURL('image/jpeg'));
    }
    InsertImages(Images);
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
