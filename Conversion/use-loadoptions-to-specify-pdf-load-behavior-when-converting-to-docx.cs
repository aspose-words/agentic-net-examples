using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToDocxConverter
{
    static void Main()
    {
        // Input PDF file.
        string pdfFilePath = @"C:\Docs\Input.pdf";

        // Output DOCX file.
        string docxFilePath = @"C:\Docs\Output.docx";

        // Configure PDF load options.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Example: do not skip images while loading the PDF.
            SkipPdfImages = false,

            // Load only the first page of the PDF (zero‑based index).
            PageIndex = 0,
            PageCount = 1
        };

        // Load the PDF into an Aspose.Words Document using the specified load options.
        Document document = new Document(pdfFilePath, loadOptions);

        // Save the loaded document as DOCX.
        document.Save(docxFilePath, SaveFormat.Docx);
    }
}
