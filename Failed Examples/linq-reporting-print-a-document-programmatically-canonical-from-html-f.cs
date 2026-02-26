// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.IO;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // HTML source that will be converted to a Word document.
        string html = "<html><body><h1>Hello Aspose.Words</h1><p>This is a paragraph.</p></body></html>";

        // Load the HTML into an Aspose.Words Document.
        // Use a MemoryStream and LoadOptions to specify the HTML format.
        using (MemoryStream htmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html)))
        {
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Document doc = new Document(htmlStream, loadOptions);

            // Rebuild the page layout so that printing uses up‑to‑date pagination.
            doc.UpdatePageLayout();

            // Create printer settings (optional – here we print all pages).
            PrinterSettings printerSettings = new PrinterSettings
            {
                PrintRange = PrintRange.AllPages
            };

            // Use AsposeWordsPrintDocument for .NET printing integration.
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);
            printDoc.PrinterSettings = printerSettings;

            // Print the document to the default printer without showing any UI.
            printDoc.Print();
        }
    }
}
