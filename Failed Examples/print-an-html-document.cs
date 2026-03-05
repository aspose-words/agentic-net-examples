// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the HTML file to be printed.
        string htmlFilePath = @"C:\Path\To\YourDocument.html";

        // Load the HTML document into an Aspose.Words Document object.
        Document doc = new Document(htmlFilePath);

        // Option 1: Print using the default printer.
        doc.Print();

        // Option 2: Print to a specific printer by name.
        // Uncomment and set the printer name if you need to target a particular printer.
        // string printerName = "Your Printer Name";
        // doc.Print(printerName);

        // Option 3: Print with custom printer settings (e.g., page range).
        // Uncomment to use advanced settings.
        // PrinterSettings settings = new PrinterSettings
        // {
        //     PrinterName = "Your Printer Name",
        //     PrintRange = PrintRange.SomePages,
        //     FromPage = 1,
        //     ToPage = doc.PageCount
        // };
        // doc.Print(settings);
    }
}
