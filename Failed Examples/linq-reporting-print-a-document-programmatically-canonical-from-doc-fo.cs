// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace AsposeWordsPrintExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOC file. The Document constructor automatically detects the format.
            Document doc = new Document("SampleDocument.doc");

            // Ensure the layout information is up‑to‑date before printing.
            doc.UpdatePageLayout();

            // -------------------------------------------------
            // 1. Print to the default printer (no UI).
            // -------------------------------------------------
            doc.Print(); // Uses the overload that prints the whole document to the default printer.

            // -------------------------------------------------
            // 2. Print a specific page range to a named printer.
            // -------------------------------------------------
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Replace with the exact name of the target printer if needed.
                // PrinterName = "My Printer Name",
                PrintRange = PrintRange.SomePages,
                FromPage = 1,
                ToPage = 2
            };

            // Directly use the Document.Print overload that accepts PrinterSettings.
            doc.Print(printerSettings);

            // -------------------------------------------------
            // 3. Advanced printing using AsposeWordsPrintDocument.
            //    This gives access to events, color mode, etc.
            // -------------------------------------------------
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc)
            {
                // Assign the same printer settings as above.
                PrinterSettings = printerSettings,

                // Example: force grayscale printing when the printer supports color.
                ColorMode = ColorPrintMode.GrayscaleAuto
            };

            // Optional: cache printer settings to reduce the first‑print latency.
            awPrintDoc.CachePrinterSettings();

            // Print the document using the AsposeWordsPrintDocument implementation.
            awPrintDoc.Print();
        }
    }
}
