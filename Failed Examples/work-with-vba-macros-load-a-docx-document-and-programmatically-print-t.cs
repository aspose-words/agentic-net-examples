// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load the DOCX document from disk using the built‑in constructor.
        Document doc = new Document("input.docx");

        // Simple case: print the whole document to the default printer.
        doc.Print();

        // ------------------------------------------------------------
        // If you need to control printer settings (e.g., select a specific printer,
        // set a page range, or use grayscale), use AsposeWordsPrintDocument.
        // ------------------------------------------------------------
        /*
        // Configure the .NET printer settings.
        PrinterSettings printerSettings = new PrinterSettings
        {
            PrinterName = "Microsoft Print to PDF", // replace with your printer name
            PrintRange = PrintRange.AllPages
        };

        // Wrap the Aspose.Words document in the AsposeWordsPrintDocument class.
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);
        printDoc.PrinterSettings = printerSettings;

        // Optional: cache printer settings to reduce the first‑print latency.
        printDoc.CachePrinterSettings();

        // Print using the specified settings.
        printDoc.Print();
        */
    }
}
