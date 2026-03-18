using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Configure HTML save options to embed fonts as Base64.
        HtmlSaveOptions options = new HtmlSaveOptions
        {
            ExportFontsAsBase64 = true,               // Embed fonts in Base64.
            CssStyleSheetType = CssStyleSheetType.Embedded, // Embed CSS directly.
            PrettyFormat = true                       // Optional: make output more readable.
        };

        // Save the document as HTML with the specified options.
        doc.Save("Output.html", options);
    }
}
