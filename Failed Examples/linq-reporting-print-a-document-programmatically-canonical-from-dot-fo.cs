// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Path to the DOT template file.
        string templatePath = @"C:\Templates\ReportTemplate.dot";

        // Load the DOT file into a Document object.
        Document doc = new Document(templatePath);

        // Print the whole document using the default printer.
        doc.Print();

        // -------------------------------------------------
        // Advanced printing: specify printer and page range.
        // -------------------------------------------------
        PrinterSettings printerSettings = new PrinterSettings
        {
            // Name of the printer to use.
            PrinterName = "Microsoft Print to PDF",

            // Print only the first two pages.
            PrintRange = PrintRange.SomePages,
            FromPage = 1,
            ToPage = 2
        };

        // Use AsposeWordsPrintDocument for richer control (optional).
        AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
        awPrintDoc.PrinterSettings = printerSettings;

        // Print using the configured settings.
        awPrintDoc.Print();

        // Alternatively, you can call the overload directly on Document:
        // doc.Print(printerSettings, "MyReport");
    }
}
