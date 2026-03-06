// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Path to the source MHTML file.
        string mhtmlPath = @"C:\Docs\Report.mhtml";

        // Load the MHTML document. The constructor automatically detects the format.
        Document doc = new Document(mhtmlPath);

        // Rebuild the page layout so that page‑related fields (PAGE, PAGES, etc.) are up‑to‑date.
        doc.UpdatePageLayout();

        // -------------------------------------------------
        // Example 1: Print to the default printer.
        // -------------------------------------------------
        // doc.Print();

        // -------------------------------------------------
        // Example 2: Print to a specific printer with a page range.
        // -------------------------------------------------
        PrinterSettings printerSettings = new PrinterSettings
        {
            // Choose the printer you want to use.
            PrinterName = "Microsoft Print to PDF",

            // Print a specific range of pages (here we print the whole document).
            PrintRange = PrintRange.SomePages,
            FromPage = 1,
            ToPage = doc.PageCount
        };

        // Use AsposeWordsPrintDocument for richer control (e.g., caching printer settings).
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);
        printDoc.PrinterSettings = printerSettings;

        // Cache printer settings to reduce the first‑print latency.
        printDoc.CachePrinterSettings();

        // Send the document to the printer.
        printDoc.Print();

        // -------------------------------------------------
        // Example 3: Directly print using Document.Print overload.
        // -------------------------------------------------
        // doc.Print(printerSettings, "MyReport");
    }
}
