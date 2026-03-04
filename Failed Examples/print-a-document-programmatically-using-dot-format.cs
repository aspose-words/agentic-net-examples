// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load an existing Word document from disk.
        // Replace the path with the actual location of your .docx file.
        Document doc = new Document(@"C:\Docs\SampleDocument.docx");

        // -----------------------------------------------------------------
        // Simple printing using the Document.Print() method.
        // This sends the whole document to the default printer.
        // -----------------------------------------------------------------
        doc.Print();

        // -----------------------------------------------------------------
        // Advanced printing using AsposeWordsPrintDocument.
        // Allows setting printer settings, color mode, and page filters.
        // -----------------------------------------------------------------
        // Create a PrinterSettings object to specify printer name and page range.
        PrinterSettings printerSettings = new PrinterSettings
        {
            // Use the name of an installed printer.
            // Example: "Microsoft Print to PDF"
            PrinterName = "Microsoft Print to PDF",
            // Print all pages (1‑based indexing).
            FromPage = 1,
            ToPage = doc.PageCount,
            PrintRange = PrintRange.SomePages
        };

        // Wrap the Aspose.Words document in a PrintDocument implementation.
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc)
        {
            PrinterSettings = printerSettings,
            // Print non‑colored pages in grayscale to save ink.
            ColorMode = ColorPrintMode.GrayscaleAuto
        };

        // Optional: cache printer settings to improve the first print call performance.
        printDoc.CachePrinterSettings();

        // Send the document to the specified printer.
        printDoc.Print();

        // After printing you can query how many pages were actually printed.
        Console.WriteLine($"Total pages printed: {printDoc.TotalPagesPrinted}");
        Console.WriteLine($"Color pages printed: {printDoc.ColorPagesPrinted}");
    }
}
