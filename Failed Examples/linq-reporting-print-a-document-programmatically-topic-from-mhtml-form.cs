// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace AsposeWordsMhtmlPrintDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the MHTML file that contains the report.
            const string mhtmlPath = @"C:\Reports\Report.mhtml";

            // Load the MHTML document. The Document constructor automatically detects the format.
            Document doc = new Document(mhtmlPath);

            // Option 1: Print directly using the Document.Print() method (default printer).
            // doc.Print();

            // Option 2: Print using a specific printer and printer settings.
            // Create a PrinterSettings object to specify printer name and page range if needed.
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Example: specify a printer installed on the system.
                // PrinterName = "Microsoft Print to PDF",

                // Example: print only pages 1 through 3.
                // PrintRange = PrintRange.SomePages,
                // FromPage = 1,
                // ToPage = 3
            };

            // Use AsposeWordsPrintDocument for more control (e.g., progress tracking, color mode).
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);
            printDoc.PrinterSettings = printerSettings;

            // Optional: set color mode (GrayscaleAuto, Color, etc.).
            // printDoc.ColorMode = ColorPrintMode.GrayscaleAuto;

            // Cache printer settings to reduce first-print latency.
            printDoc.CachePrinterSettings();

            // Print the document.
            printDoc.Print();

            // If you prefer the simpler overload, you can also call:
            // doc.Print(printerSettings);
        }
    }
}
