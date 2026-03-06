// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using System.Drawing.Printing;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add some text to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This document will be printed.");

        // Print the document using the default printer.
        doc.Print();

        // Uncomment the following lines to print to a specific printer by name.
        // string printerName = PrinterSettings.InstalledPrinters[0];
        // doc.Print(printerName);
    }
}
