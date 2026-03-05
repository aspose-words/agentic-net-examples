// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the DOC file you want to print.
        string docPath = @"C:\Docs\SampleDocument.doc";

        // Load the existing DOC document using the provided constructor.
        Document doc = new Document(docPath);

        // Option 1: Print to the default printer.
        doc.Print();

        // Option 2: Print to a specific printer (uncomment and set the printer name if needed).
        // string printerName = "Your Printer Name";
        // doc.Print(printerName);

        // Option 3: Print with custom printer settings (e.g., print only pages 1‑2).
        // PrinterSettings settings = new PrinterSettings
        // {
        //     PrintRange = PrintRange.SomePages,
        //     FromPage = 1,
        //     ToPage = 2,
        //     PrinterName = "Your Printer Name"
        // };
        // doc.Print(settings);
    }
}
