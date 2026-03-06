// ALL ATTEMPTS FAILED. Below is the last generated code.

using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document from a file.
        Document doc = new Document("MyDocument.docx");

        // Print the whole document using the default printer.
        doc.Print();

        // Example: print to a specific printer by name.
        // string printerName = PrinterSettings.InstalledPrinters[0];
        // doc.Print(printerName);
    }
}
