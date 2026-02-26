// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Path to the source DOC file.
        const string inputPath = @"C:\Docs\Sample.doc";

        // Load the DOC document. The constructor automatically detects the format.
        Document doc = new Document(inputPath);

        // Ensure the layout information is up‑to‑date before printing.
        doc.UpdatePageLayout();

        // -----------------------------------------------------------------
        // Option 1: Simple print using the default printer.
        // -----------------------------------------------------------------
        doc.Print(); // Prints the whole document to the default printer.

        // -----------------------------------------------------------------
        // Option 2: Print with explicit printer settings using AsposeWordsPrintDocument.
        // -----------------------------------------------------------------
        // Create a PrintDocument that wraps the Aspose.Words document.
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);

        // Configure printer settings (e.g., select a specific printer and page range).
        PrinterSettings printerSettings = new PrinterSettings
        {
            // Replace with the name of an installed printer if needed.
            // PrinterName = "Your Printer Name",
            PrintRange = PrintRange.SomePages,
            FromPage = 1,
            ToPage = doc.PageCount // print all pages; adjust as required
        };
        printDoc.PrinterSettings = printerSettings;

        // Optional: reduce the first‑print overhead.
        printDoc.CachePrinterSettings();

        // Print the document using the configured settings.
        printDoc.Print();
    }
}
