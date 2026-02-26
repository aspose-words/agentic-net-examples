// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load the Word template in DOT format.
        Document doc = new Document("Template.dot");

        // Print the whole document to the default printer.
        doc.Print();

        // -------------------------------------------------
        // Advanced printing: specify printer and page range.
        // -------------------------------------------------
        PrinterSettings printerSettings = new PrinterSettings
        {
            // Replace with the desired printer name.
            PrinterName = "Microsoft Print to PDF",
            PrintRange = PrintRange.SomePages,
            FromPage = 1,
            ToPage = doc.PageCount // print all pages
        };

        // Use AsposeWordsPrintDocument for fine‑grained control.
        AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
        awPrintDoc.PrinterSettings = printerSettings;

        // Print without showing any UI.
        awPrintDoc.Print();
    }
}
