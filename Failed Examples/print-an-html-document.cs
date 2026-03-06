// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class PrintHtmlExample
{
    static void Main()
    {
        // Path to the HTML file to be printed.
        const string htmlFilePath = @"C:\Docs\sample.html";

        // Load the HTML document into an Aspose.Words Document object.
        Document doc = new Document(htmlFilePath);

        // Option 1: Print directly using the default printer.
        doc.Print();

        // Option 2: Print using a specific printer with custom settings.
        // Uncomment the following lines if you need to target a particular printer.
        /*
        PrinterSettings printerSettings = new PrinterSettings
        {
            PrinterName = "Microsoft Print to PDF", // replace with your printer name
            PrintRange = PrintRange.AllPages
        };

        // Use AsposeWordsPrintDocument for advanced control (e.g., color mode, page filtering).
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);
        printDoc.PrinterSettings = printerSettings;
        // Example: print in grayscale.
        // printDoc.ColorMode = ColorPrintMode.GrayscaleAuto;
        printDoc.Print();
        */
    }
}
