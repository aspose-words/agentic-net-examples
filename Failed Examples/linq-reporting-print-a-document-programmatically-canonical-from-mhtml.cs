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

        // Rebuild the page layout so that pagination information is current.
        doc.UpdatePageLayout();

        // -----------------------------------------------------------------
        // Simple printing: send the document to the default printer.
        // -----------------------------------------------------------------
        doc.Print();

        // -----------------------------------------------------------------
        // Advanced printing: use AsposeWordsPrintDocument to specify printer
        // settings, cache printer data, and print to a chosen printer.
        // -----------------------------------------------------------------
        PrinterSettings printerSettings = new PrinterSettings
        {
            // Replace with the name of the printer you want to use.
            PrinterName = "Microsoft Print to PDF",
            PrintRange = PrintRange.AllPages
        };

        // Wrap the Document in AsposeWordsPrintDocument.
        AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
        awPrintDoc.PrinterSettings = printerSettings;

        // Cache printer settings to improve the first print call performance.
        awPrintDoc.CachePrinterSettings();

        // Print the document using the specified settings.
        awPrintDoc.Print();
    }
}
