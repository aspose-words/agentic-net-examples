using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToHtmlWithExternalCss
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Docs\source.pdf";

        // Path where the resulting HTML file will be saved.
        string htmlPath = @"C:\Docs\output.html";

        // Load the PDF document.
        Document doc = new Document(pdfPath);

        // Configure HTML fixed save options.
        // Setting ExportEmbeddedCss to true creates a separate CSS file
        // and links it from the generated HTML document.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            ExportEmbeddedCss = true
        };

        // Save the document as HTML using the configured options.
        doc.Save(htmlPath, htmlOptions);
    }
}
