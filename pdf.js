pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@2.8.335/build/pdf.worker.js';

const file = document.getElementById("file");
const UploadButton = document.getElementById("upload");
const canvas = document.getElementById('the-canvas');
const fileReader = new FileReader();

fileReader.onload = function (e) {
    readPDFFile(new Uint8Array(e.target.result));
};


async function readPDFFile(pdf_data) {
    const pdf = await pdfjsLib.getDocument({ data: pdf_data }).promise;

    const page_count = pdf.numPages;
    for (let i = 1; i <= page_count; i++) {

        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: 1.5 });


        const context = canvas.getContext('2d');
        canvas.height = viewport.height;
        canvas.width = viewport.width;


        await page.render({ canvasContext: context, viewport: viewport }).promise
        InsertImages(canvas.toDataURL('image/jpeg'));
    }
    file.value = "";
}

file.onchange = function (e) {
    let PDFFile = file.files[0];
    if (!PDFFile) return;

    if (PDFFile.type !== "application/pdf") {
        file.value = "";
        alert("Please select a PDF file.");
        return;
    }
};

UploadButton.onclick = function () {
    let PDFFile = file.files[0];
    if (!PDFFile) {
        alert("Please select a PDF file.");
    };

    fileReader.readAsArrayBuffer(PDFFile);
};

const InsertImages = (base64Image) => {
    // Run a batch operation against the Word object model.
    Word.run(async function (context) {

        const body = context.document.body;

        base64Image = base64Image.replace(/^data:image\/\w+;base64,/, "");

        body.insertInlinePictureFromBase64(base64Image, "End");

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        await context.sync();
        console.log('Added base64 encoded text to the beginning of the document body.');
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
