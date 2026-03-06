using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class PdfConversionExample
{
    static void Main()
    {
        // Input PDF file.
        string inputPdfPath = @"C:\Docs\Input.pdf";

        // Output files in various formats.
        string outputDocxPath = @"C:\Docs\Output.docx";
        string outputHtmlPath = @"C:\Docs\Output.html";
        string outputTxtPath  = @"C:\Docs\Output.txt";

        // Load the PDF document. The constructor automatically detects the format.
        Document pdfDocument = new Document(inputPdfPath);

        // Convert PDF to DOCX.
        pdfDocument.Save(outputDocxPath, SaveFormat.Docx);

        // Convert PDF to HTML.
        pdfDocument.Save(outputHtmlPath, SaveFormat.Html);

        // Convert PDF to plain text.
        pdfDocument.Save(outputTxtPath, SaveFormat.Text);

        // ------------------------------------------------------------
        // Example: Load PDF while skipping images, then save to DOCX.
        // ------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Do not load images from the PDF.
            SkipPdfImages = true
        };

        // Load the PDF with the specified options.
        Document pdfWithoutImages = new Document(inputPdfPath, loadOptions);

        // Save the image‑less document to DOCX.
        string outputDocxNoImagesPath = @"C:\Docs\Output_NoImages.docx";
        pdfWithoutImages.Save(outputDocxNoImagesPath, SaveFormat.Docx);
    }
}
