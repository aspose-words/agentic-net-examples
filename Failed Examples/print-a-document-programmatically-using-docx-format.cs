// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Print the document to the default printer.
        doc.Print();

        // Uncomment the following lines to print to a specific printer.
        // PrinterSettings printerSettings = new PrinterSettings();
        // printerSettings.PrinterName = "YourPrinterName";
        // doc.Print(printerSettings);
    }
}
