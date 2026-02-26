// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace AsposeWordsRtfPrintDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source RTF file.
            const string rtfPath = @"C:\Docs\SampleDocument.rtf";

            // Load the RTF document. The constructor automatically detects the format.
            Document doc = new Document(rtfPath);

            // Ensure the layout information is up‑to‑date before printing.
            doc.UpdatePageLayout();

            // -------------------------------------------------
            // 1) Print to the default printer (no UI).
            // -------------------------------------------------
            doc.Print();

            // -------------------------------------------------
            // 2) Print to a specific printer with a custom page range.
            // -------------------------------------------------
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Replace with the exact name of the target printer.
                PrinterName = "Microsoft Print to PDF",

                // Example: print only pages 2 through 4 (1‑based indexing).
                PrintRange = PrintRange.SomePages,
                FromPage = 2,
                ToPage = 4
            };

            // Use Aspose.Words' built‑in Print method that accepts PrinterSettings.
            doc.Print(printerSettings);

            // -------------------------------------------------
            // 3) Advanced printing using AsposeWordsPrintDocument.
            //    This gives access to events such as page‑by‑page tracking.
            // -------------------------------------------------
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc)
            {
                PrinterSettings = printerSettings,
                // Example: force grayscale output if the printer supports it.
                ColorMode = ColorPrintMode.GrayscaleAuto
            };

            // Optional: cache printer settings to reduce the first‑print latency.
            awPrintDoc.CachePrinterSettings();

            // Print the document using the configured AsposeWordsPrintDocument.
            awPrintDoc.Print();

            Console.WriteLine("Printing completed.");
        }
    }
}
