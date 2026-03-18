using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToHtmlWithEmbeddedFonts
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Docs\source.pdf";

        // Path where the resulting HTML file will be saved.
        string htmlPath = @"C:\Docs\output.html";

        // Load the PDF document.
        Document doc = new Document(pdfPath);

        // Configure HTML fixed save options to embed fonts as Base64.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            // When true, font data is embedded directly into the CSS as Base64 strings.
            ExportEmbeddedFonts = true,

            // Optional: embed CSS into the HTML to keep everything in a single file.
            ExportEmbeddedCss = true,

            // Optional: embed images as Base64 as well.
            ExportEmbeddedImages = true
        };

        // Save the document as HTML using the configured options.
        doc.Save(htmlPath, htmlOptions);
    }
}
