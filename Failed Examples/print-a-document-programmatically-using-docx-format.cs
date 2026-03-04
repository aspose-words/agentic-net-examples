// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();                     // Document() constructor (create rule)
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");                  // Add some content.

        // Save the document as a DOCX file (demonstrates the Save(string) rule).
        string docPath = "HelloWorld.docx";
        doc.Save(docPath);                                 // Save(string) method

        // Print the document to the default printer.
        doc.Print();                                       // Print() method (default printer)

        // Uncomment the following lines to print to a specific printer by name.
        // string printerName = new PrinterSettings().InstalledPrinters[0];
        // doc.Print(printerName);                         // Print(string) overload
    }
}
