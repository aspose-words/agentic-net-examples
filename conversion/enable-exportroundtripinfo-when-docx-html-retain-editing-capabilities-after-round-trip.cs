using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ExportRoundTripInfoExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document and add some sample content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document created at runtime.");
            builder.InsertParagraph();
            builder.Writeln("It will be saved to HTML with round‑trip information enabled.");

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Enable round‑trip information so that the HTML can be loaded back into Aspose.Words
                // with full fidelity (preserves headers/footers, comments, tab stops, etc.).
                ExportRoundtripInformation = true
            };

            // Save the document as HTML using the configured options.
            doc.Save("OutputDocument.html", htmlOptions);
        }
    }
}
