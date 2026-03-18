using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class HtmlToPdfConverter
{
    static void Main()
    {
        // Input HTML file that may contain external JavaScript.
        string htmlPath = @"C:\Input\sample.html";

        // Desired output PDF file.
        string pdfPath = @"C:\Output\sample.pdf";

        // Load options: ignore <noscript> elements so script‑related content is not rendered.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        loadOptions.IgnoreNoscriptElements = true;

        // Load the HTML document with the specified options.
        Document doc = new Document(htmlPath, loadOptions);

        // PDF save options (default configuration is sufficient for static rendering).
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Save the loaded document as PDF.
        doc.Save(pdfPath, saveOptions);
    }
}
