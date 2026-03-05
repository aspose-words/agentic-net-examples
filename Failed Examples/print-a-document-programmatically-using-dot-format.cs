// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Rendering;
using System.Drawing.Printing;

class Program
{
    static void Main()
    {
        // Load the Word document from a file.
        // Replace the path with the actual location of your document.
        Document doc = new Document(@"C:\Docs\SampleDocument.docx");

        // Option 1: Print directly using the Document.Print() method.
        // This sends the whole document to the default printer.
        doc.Print();

        // Option 2: Use AsposeWordsPrintDocument for more advanced scenarios
        // (e.g., specifying printer settings, color mode, or page filters).
        // Uncomment the following block if you need such control.

        /*
        // Create a PrintDocument wrapper around the Aspose.Words document.
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);

        // Configure printer settings (example: print to a specific printer).
        PrinterSettings printerSettings = new PrinterSettings
        {
            PrinterName = "Microsoft Print to PDF", // replace with your printer name
            FromPage = 1,
            ToPage = doc.PageCount,
            PrintRange = PrintRange.AllPages
        };
        printDoc.PrinterSettings = printerSettings;

        // Optional: set color mode (e.g., print non‑color pages in grayscale).
        printDoc.ColorMode = ColorPrintMode.GrayscaleAuto;

        // Cache printer settings to improve the first print call performance.
        printDoc.CachePrinterSettings();

        // Print the document using the configured settings.
        printDoc.Print();
        */

        // End of program.
    }
}
