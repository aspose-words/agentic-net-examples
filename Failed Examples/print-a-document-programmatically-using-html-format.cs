// ALL ATTEMPTS FAILED. Below is the last generated code.

using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load an HTML document from a file.
        // The Document constructor automatically detects the HTML format.
        Document doc = new Document("input.html");

        // Print the whole document to the default printer.
        doc.Print();

        // If you need to print to a specific printer, uncomment the following lines:
        // PrinterSettings printerSettings = new PrinterSettings();
        // printerSettings.PrinterName = "YourPrinterName";
        // doc.Print(printerSettings);
    }
}
