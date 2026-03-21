using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportHtmlWithDynamicColors
{
    static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, world! This is a sample document with dynamic colors.");

        // Configure HTML save options.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportRoundtripInformation = true, // Preserve Aspose.Words round‑trip info (e.g., dynamic colors).
            PrettyFormat = true,                // Make the generated HTML easier to read.
            Encoding = new UTF8Encoding(false) // UTF‑8 without BOM.
        };

        // Save the document as HTML using the configured options.
        doc.Save("Output.html", options);
        Console.WriteLine("HTML document saved successfully.");
    }
}
