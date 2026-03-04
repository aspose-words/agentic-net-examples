// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load an existing DOC file from disk.
        Document doc = new Document("Input.doc");

        // Ensure the layout information is up‑to‑date before printing.
        doc.UpdatePageLayout();

        // 1. Print the document to the default printer.
        doc.Print();

        // 2. Print the document to a specific printer using PrinterSettings.
        PrinterSettings printerSettings = new PrinterSettings
        {
            // Replace with the exact name of the target printer.
            PrinterName = "YourPrinterName",
            PrintRange = PrintRange.AllPages
        };
        doc.Print(printerSettings);

        // 3. Use AsposeWordsPrintDocument for more control (e.g., preview, progress tracking).
        AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
        awPrintDoc.PrinterSettings = printerSettings;
        awPrintDoc.Print();
    }
}
