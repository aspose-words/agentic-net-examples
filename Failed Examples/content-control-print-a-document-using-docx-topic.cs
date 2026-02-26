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
            // Load an existing DOCX file. The Document constructor handles the loading lifecycle.
            Document doc = new Document(@"C:\Docs\SampleDocument.docx");

            // -------------------------------------------------
            // Example 1: Print the document using the default printer.
            // -------------------------------------------------
            // The Print() method prints the whole document to the default printer.
            doc.Print();

            // -------------------------------------------------
            // Example 2: Print the document to a specific printer with custom settings.
            // -------------------------------------------------
            // Create a PrinterSettings object to control the print job.
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Choose a printer by name (replace with an installed printer on your system).
                PrinterName = "Microsoft Print to PDF",

                // Print only pages 1 through 3 (page indexing is 1‑based).
                PrintRange = PrintRange.SomePages,
                FromPage = 1,
                ToPage = 3
            };

            // Print using the specified printer settings.
            doc.Print(printerSettings);

            // -------------------------------------------------
            // Example 3: Print using a named printer and assign a document name.
            // -------------------------------------------------
            // The overload that accepts a printer name prints without UI.
            string printerName = printerSettings.PrinterName;
            doc.Print(printerName);

            // -------------------------------------------------
            // Example 4: Use AsposeWordsPrintDocument for advanced tracking.
            // -------------------------------------------------
            // Wrap the Document in AsposeWordsPrintDocument to access progress events.
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
            awPrintDoc.PrinterSettings = printerSettings;

            // Optional: cache printer settings to reduce first‑print latency.
            awPrintDoc.CachePrinterSettings();

            // Print the document via the AsposeWordsPrintDocument instance.
            awPrintDoc.Print();

            // Output the number of pages printed (useful for logging).
            Console.WriteLine($"Total pages printed: {awPrintDoc.TotalPagesPrinted}");
        }
    }
}
