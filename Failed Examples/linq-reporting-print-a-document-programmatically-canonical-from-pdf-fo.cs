// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Path to the PDF file that will be printed.
        string pdfPath = @"C:\Docs\Sample.pdf";

        // Load the PDF into an Aspose.Words Document.
        Document doc = new Document(pdfPath);

        // Rebuild the page layout to ensure accurate pagination before printing.
        doc.UpdatePageLayout();

        // -----------------------------------------------------------------
        // Simple printing: send the whole document to the default printer.
        // -----------------------------------------------------------------
        doc.Print();

        // -----------------------------------------------------------------
        // Advanced printing: specify printer, page range, and use the
        // AsposeWordsPrintDocument wrapper for richer control.
        // -----------------------------------------------------------------
        PrinterSettings printerSettings = new PrinterSettings
        {
            // Replace with the actual printer name you want to use.
            PrinterName = "Microsoft Print to PDF",
            PrintRange = PrintRange.SomePages,
            FromPage = 1,
            ToPage = 2
        };

        // Create the Aspose.Words implementation of PrintDocument.
        AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
        awPrintDoc.PrinterSettings = printerSettings;

        // Cache printer settings to reduce the first‑call overhead.
        awPrintDoc.CachePrinterSettings();

        // Print the document using the specified settings.
        awPrintDoc.Print();
    }
}
