pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@2.8.335/build/pdf.worker.js';

const file = document.getElementById("file");
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
}

file.onchange = function (e) {
    let PDFFile = file.files[0];
    if (!PDFFile) return;

    if (PDFFile.type !== "application/pdf") {
        file.value = "";
        alert("Please select a PDF file.");
        return;
    }

    fileReader.readAsArrayBuffer(PDFFile);
};

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = getDocumentAsCompressed;
    }
});

const InsertImages = (base64Image) => {
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        const body = context.document.body;

        body.insertInlinePictureFromBase64(base64Image, "After");

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log('Added base64 encoded text to the beginning of the document body.');
        });
    })
        .catch(function (error) {
            console.log('Error: ' + JSON.stringify(error));
            if (error instanceof OfficeExtension.Error) {
                console.log('Debug info: ' + JSON.stringify(error.debugInfo));
            }
        });
}
