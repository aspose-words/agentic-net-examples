using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportHtmlWithDynamicColors
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("Input.docx");

        // Configure HTML save options.
        // ExportRoundtripInformation preserves Aspose.Words specific CSS (‑aw‑*) that stores
        // original formatting such as dynamic colors applied via tags.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportRoundtripInformation = true,   // keep round‑trip information for colors, etc.
            PrettyFormat = true,                // makes the generated HTML easier to read.
            Encoding = new UTF8Encoding(false) // UTF‑8 without BOM.
        };

        // Save the document as HTML using the configured options.
        doc.Save("Output.html", options);
    }
}
