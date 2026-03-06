// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load the DOCX document from the file system.
        Document doc = new Document("input.docx");

        // Set up printer settings (optional – you can change the printer name as needed).
        PrinterSettings printerSettings = new PrinterSettings
        {
            PrinterName = "Microsoft Print to PDF",
            PrintRange = PrintRange.AllPages
        };

        // Create the Aspose.Words print document that wraps the loaded Document.
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);
        printDoc.PrinterSettings = printerSettings;

        // Cache printer settings to reduce the time of the first print call.
        printDoc.CachePrinterSettings();

        // Print the document.
        printDoc.Print();

        // Display the total number of pages that were printed.
        Console.WriteLine($"Total pages printed: {printDoc.TotalPagesPrinted}");
    }
}
