using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ExportRoundTripInfoExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCX document.
            Document doc = new Document("InputDocument.docx");

            // Create HtmlSaveOptions for HTML output.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Enable round‑trip information so that the HTML can be loaded back into Aspose.Words
            // with full fidelity (preserves headers/footers, comments, tab stops, etc.).
            htmlOptions.ExportRoundtripInformation = true;

            // Save the document as HTML using the configured options.
            doc.Save("OutputDocument.html", htmlOptions);
        }
    }
}
