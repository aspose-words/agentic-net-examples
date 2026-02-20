using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        string docPath = @"C:\Input\sample.doc";

        // Path where the resulting HTML file will be saved.
        string htmlPath = @"C:\Output\sample.html";

        // Load the DOC document.
        Document doc = new Document(docPath);

        // Configure HTML save options.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Use HTML5 standard for the output.
            HtmlVersion = HtmlVersion.Html5,
            // Produce nicely indented (pretty) HTML.
            PrettyFormat = true,
            // Export the document's headers and footers as part of the body.
            ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection,
            // Export images as separate files (default behavior).
            ExportImagesAsBase64 = false
        };

        // Save the document as HTML.
        doc.Save(htmlPath, saveOptions);
    }
}
