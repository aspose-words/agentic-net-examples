// ALL ATTEMPTS FAILED. Below is the last generated code.

using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world!");

        // Print the whole document to the default printer.
        doc.Print();

        // Uncomment the following lines to print to a specific printer by name.
        // string printerName = PrinterSettings.InstalledPrinters[0];
        // doc.Print(printerName);
    }
}
