// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document from the file system.
        // This uses the Document(string) constructor – the standard load rule.
        Document doc = new Document("Input.docx");

        // -----------------------------------------------------------------
        // 1. Print the whole document to the default printer.
        //    The parameter‑less Print() method follows the built‑in print rule.
        // -----------------------------------------------------------------
        doc.Print();

        // -----------------------------------------------------------------
        // 2. Print the document to a specific printer with custom settings.
        //    Here we create a PrinterSettings object, configure it, and pass it
        //    to the Print(PrinterSettings, string) overload – the standard
        //    printing rule that accepts printer settings and a document name.
        // -----------------------------------------------------------------
        PrinterSettings printerSettings = new PrinterSettings();

        // Set the target printer (replace with an installed printer name).
        printerSettings.PrinterName = "MyPrinter";

        // Example: print only the first three pages.
        printerSettings.PrintRange = PrintRange.SomePages;
        printerSettings.FromPage = 1;               // 1‑based page index
        printerSettings.ToPage = Math.Min(3, doc.PageCount);

        // Print using the configured settings and give the job a friendly name.
        doc.Print(printerSettings, "MyPrintedDocument");
    }
}
