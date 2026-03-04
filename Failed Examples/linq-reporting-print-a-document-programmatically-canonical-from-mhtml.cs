// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsMhtmlPrintDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source MHTML file.
            const string mhtmlFilePath = @"C:\Docs\Report.mhtml";

            // -----------------------------------------------------------------
            // 1. Load the MHTML document.
            // -----------------------------------------------------------------
            // The Document constructor automatically detects the format from the file extension.
            Document doc = new Document(mhtmlFilePath);

            // Ensure the layout information is up‑to‑date before printing.
            doc.UpdatePageLayout();

            // -----------------------------------------------------------------
            // 2. Print the document to the default printer.
            // -----------------------------------------------------------------
            doc.Print();

            // -----------------------------------------------------------------
            // 3. (Optional) Print to a specific printer with custom settings.
            // -----------------------------------------------------------------
            PrinterSettings printerSettings = new PrinterSettings
            {
                // Replace with the exact name of the target printer.
                PrinterName = "My Printer",
                // Example: print only the first three pages.
                PrintRange = PrintRange.SomePages,
                FromPage = 1,
                ToPage = 3
            };

            // Print using the supplied PrinterSettings.
            doc.Print(printerSettings);

            // -----------------------------------------------------------------
            // 4. (Optional) Save the document as PDF – demonstrates the required
            //    save lifecycle rule.
            // -----------------------------------------------------------------
            const string pdfOutputPath = @"C:\Docs\Report.pdf";
            doc.Save(pdfOutputPath, SaveFormat.Pdf);
        }
    }
}
