using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToHtmlConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Docs\source.pdf";

        // Path where the resulting HTML file will be saved.
        string htmlPath = @"C:\Docs\result.html";

        // Load the PDF document. No special load options are required for image extraction.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document document = new Document(pdfPath, loadOptions);

        // Configure HTML save options to embed images as Base64 data URIs.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportImagesAsBase64 = true,   // Embed images directly in the HTML.
            PrettyFormat = true            // Optional: make the output HTML more readable.
        };

        // Save the document as HTML with embedded images.
        document.Save(htmlPath, htmlOptions);
    }
}
