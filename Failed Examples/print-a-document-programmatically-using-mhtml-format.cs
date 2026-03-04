// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the MHTML file that we want to print.
        const string mhtmlPath = @"C:\Docs\Sample.mht";

        // Load the MHTML document into an Aspose.Words Document object.
        Document doc = new Document(mhtmlPath);

        // Option 1: Print using the default printer.
        doc.Print();

        // Option 2: Print using a specific printer by name.
        // Uncomment and set the printer name if you need to target a particular printer.
        // string printerName = "Your Printer Name";
        // doc.Print(printerName);

        // Option 3: Print with detailed printer settings (e.g., page range).
        // Uncomment to use custom settings.
        // PrinterSettings settings = new PrinterSettings
        // {
        //     PrinterName = "Your Printer Name",
        //     PrintRange = PrintRange.SomePages,
        //     FromPage = 1,
        //     ToPage = 2
        // };
        // doc.Print(settings);
    }
}
