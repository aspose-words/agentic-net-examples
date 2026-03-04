using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

class PdfConversion
{
    static void Main()
    {
        // Input PDF file.
        string inputPdfPath = @"C:\Docs\Input.pdf";

        // Output files in different formats.
        string outputDocxPath = @"C:\Docs\Converted.docx";
        string outputHtmlPath = @"C:\Docs\Converted.html";
        string outputJpegPath = @"C:\Docs\Page1.jpg";

        // -----------------------------------------------------------------
        // Load the PDF with custom load options.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            // Example: do not skip images (set true to skip them).
            SkipPdfImages = false,
            // Load all pages starting from the first page.
            PageIndex = 0,
            PageCount = int.MaxValue
        };

        Document pdfDocument = new Document(inputPdfPath, loadOptions);

        // -----------------------------------------------------------------
        // Convert PDF to DOCX with password protection.
        // -----------------------------------------------------------------
        OoxmlSaveOptions docxSaveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "Secret123"
        };

        pdfDocument.Save(outputDocxPath, docxSaveOptions);

        // -----------------------------------------------------------------
        // Convert PDF to HTML (no custom options required for basic HTML).
        // -----------------------------------------------------------------
        pdfDocument.Save(outputHtmlPath, SaveFormat.Html);

        // -----------------------------------------------------------------
        // Render the first page of the PDF to a JPEG image with custom options.
        // -----------------------------------------------------------------
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Render only the first page (zero‑based index).
            PageSet = new PageSet(0),
            // Set JPEG quality (0‑100).
            JpegQuality = 90,
            // Enable high‑quality rendering for better visual fidelity.
            UseHighQualityRendering = true
        };

        pdfDocument.Save(outputJpegPath, jpegOptions);
    }
}
