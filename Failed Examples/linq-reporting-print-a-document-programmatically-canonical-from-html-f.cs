// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace AsposeWordsHtmlPrintExample
{
    class Program
    {
        static void Main()
        {
            // Path to the HTML file that will be loaded.
            // Aspose.Words automatically detects the format from the file extension.
            string htmlFilePath = @"C:\Docs\Report.html";

            // Load the HTML document into an Aspose.Words Document object.
            // This uses the Document(string) constructor – the standard creation rule.
            Document doc = new Document(htmlFilePath);

            // Ensure the layout information is up‑to‑date before printing.
            // This call is optional because the first print operation triggers layout,
            // but it is shown here for completeness.
            doc.UpdatePageLayout();

            // Print the whole document to the default printer.
            // Uses the Document.Print() method – the standard printing rule.
            doc.Print();

            // ------------------------------------------------------------
            // Optional: print with custom printer settings using AsposeWordsPrintDocument.
            // ------------------------------------------------------------
            // Create a PrinterSettings object to specify printer name and page range.
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Example: select a specific printer by name.
                // PrinterName = "Your Printer Name",

                // Print only the first two pages (1‑based indexing).
                PrintRange = PrintRange.SomePages,
                FromPage = 1,
                ToPage = 2
            };

            // Wrap the Document in AsposeWordsPrintDocument to apply the settings.
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc)
            {
                PrinterSettings = printerSettings,
                // Example: print in grayscale.
                ColorMode = ColorPrintMode.GrayscaleAuto
            };

            // Cache printer settings to reduce the first‑print overhead.
            printDoc.CachePrinterSettings();

            // Perform the actual printing.
            printDoc.Print();
        }
    }
}
