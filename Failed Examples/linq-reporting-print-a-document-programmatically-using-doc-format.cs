// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

namespace AsposeWordsPrintExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOC document from disk.
            // The constructor Document(string) opens the file and detects the format automatically.
            Document doc = new Document("InputDocument.doc");

            // Option 1: Print to the default printer.
            doc.Print();

            // Option 2: Print to a specific printer by name.
            // Uncomment the following lines and replace with an actual printer name if needed.
            // string printerName = "Your Printer Name";
            // doc.Print(printerName);

            // Option 3: Print with custom printer settings (e.g., print only a range of pages).
            // Uncomment to use.
            // PrinterSettings settings = new PrinterSettings
            // {
            //     PrinterName = "Your Printer Name",
            //     PrintRange = PrintRange.SomePages,
            //     FromPage = 1,
            //     ToPage = 2
            // };
            // doc.Print(settings);
        }
    }
}
