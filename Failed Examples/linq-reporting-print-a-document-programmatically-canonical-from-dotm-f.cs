// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace AsposeWordsPrintDemo
{
    class Program
    {
        static void Main()
        {
            // Load the DOTM template from file system.
            // The Document constructor handles loading and determines the format automatically.
            Document doc = new Document("Template.dotm");

            // Ensure the layout is up‑to‑date before printing.
            doc.UpdatePageLayout();

            // Option 1: Print to the default printer.
            doc.Print();

            // Option 2: Print to a specific printer with custom settings.
            // Uncomment the following lines to use a named printer and a page range.
            /*
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Replace with the exact name of the target printer.
                PrinterName = "Your Printer Name",
                // Print only the first three pages.
                PrintRange = PrintRange.SomePages,
                FromPage = 1,
                ToPage = 3
            };

            // Use Aspose.Words' built‑in Print method that accepts PrinterSettings.
            doc.Print(printerSettings);
            */

            // Alternatively, use the AsposeWordsPrintDocument wrapper for more advanced scenarios
            // such as progress tracking or print preview.
            /*
            AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);
            printDoc.PrinterSettings = printerSettings;
            printDoc.Print();
            */
        }
    }
}
