// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using System.Drawing.Printing;

class Program
{
    static void Main()
    {
        // Load the DOTX template file.
        Document doc = new Document("Template.dotx");

        // Optionally add content to the document before printing.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document was printed from a DOTX template.");

        // Print the document using the default printer.
        doc.Print();

        // Uncomment the following lines to print to a specific printer.
        // string printerName = PrinterSettings.InstalledPrinters[0];
        // doc.Print(printerName);
    }
}
