using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToHtmlConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Path to the output HTML file.
        string htmlPath = @"C:\Output\sample.html";

        // Load the PDF document using PdfLoadOptions.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document document = new Document(pdfPath, loadOptions);

        // Configure HtmlSaveOptions to embed images as Base64.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true, // Embed images directly in the HTML.
            PrettyFormat = true          // Optional: make the HTML output more readable.
        };

        // Save the document as HTML with embedded images.
        document.Save(htmlPath, saveOptions);
    }
}
