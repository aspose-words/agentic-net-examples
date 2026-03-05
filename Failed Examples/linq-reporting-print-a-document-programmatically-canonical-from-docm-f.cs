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
            // Path to the source DOCM file.
            const string docmPath = @"C:\Docs\Template.docm";

            // Load the DOCM document. The Document constructor automatically detects the format.
            Document doc = new Document(docmPath);

            // Ensure the layout is up‑to‑date before printing.
            doc.UpdatePageLayout();

            // Optionally update fields (e.g., DATE, PAGE) so the printed output reflects current data.
            doc.UpdateFields();

            // Create a PrinterSettings instance to control the print job.
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Print all pages; change to PrintRange.SomePages to limit the range.
                PrintRange = PrintRange.AllPages,
                // Example: select a specific printer by name.
                // PrinterName = "Microsoft Print to PDF"
            };

            // Use Aspose.Words' implementation of PrintDocument for better Word‑specific handling.
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc)
            {
                PrinterSettings = printerSettings,
                // Example: print in grayscale.
                ColorMode = ColorPrintMode.GrayscaleAuto
            };

            // Cache printer settings to reduce the first‑print latency.
            printDoc.CachePrinterSettings();

            // Print the document.
            printDoc.Print();

            // Optional: output the number of pages printed in color (should be 0 for grayscale).
            Console.WriteLine($"Color pages printed: {printDoc.ColorPagesPrinted}");
        }
    }
}
