// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace AsposeWordsPrintingDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source RTF file.
            const string rtfPath = @"C:\Docs\SampleDocument.rtf";

            // Load the RTF document using the Document constructor (lifecycle rule: load).
            Document doc = new Document(rtfPath);

            // Ensure the page layout is up‑to‑date before printing.
            // This is required when the document has been modified after the first render.
            doc.UpdatePageLayout();

            // -----------------------------------------------------------------
            // Option 1: Print to the default printer (canonical simple approach).
            // -----------------------------------------------------------------
            doc.Print(); // Uses Document.Print() overload – prints the whole document.

            // -----------------------------------------------------------------
            // Option 2: Print to a specific printer with custom settings.
            // -----------------------------------------------------------------
            // Create a PrinterSettings object to define the printer and page range.
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Example: select a printer by name (replace with an installed printer).
                PrinterName = "Microsoft Print to PDF",
                // Print only the first two pages.
                PrintRange = PrintRange.SomePages,
                FromPage = 1,
                ToPage = 2
            };

            // Use AsposeWordsPrintDocument for richer control (e.g., progress tracking).
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc)
            {
                PrinterSettings = printerSettings,
                // Example: force grayscale printing.
                ColorMode = ColorPrintMode.GrayscaleAuto
            };

            // Optional: cache printer settings to reduce first‑print latency.
            printDoc.CachePrinterSettings();

            // Print using the configured settings.
            printDoc.Print();

            // -----------------------------------------------------------------
            // End of demo.
            // -----------------------------------------------------------------
        }
    }
}
