// ALL ATTEMPTS FAILED. Below is the last generated code.

using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOTX template file.
        // The Document constructor automatically detects the format from the file extension.
        Document doc = new Document("Template.dotx");

        // Print the whole document to the default printer.
        doc.Print();

        // If you need to print to a specific printer, uncomment the following lines:
        // string printerName = PrinterSettings.InstalledPrinters[0]; // Choose a printer from the installed list.
        // doc.Print(printerName);
    }
}
