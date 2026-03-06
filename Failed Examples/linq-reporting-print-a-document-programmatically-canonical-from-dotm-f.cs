// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOTM template from disk.
        // The Document(string) constructor automatically detects the format.
        Document doc = new Document("Template.dotm");

        // Rebuild the page layout to ensure that pagination and fields are up‑to‑date
        // before sending the document to the printer.
        doc.UpdatePageLayout();

        // Print the whole document to the default printer.
        doc.Print();

        // ------------------------------------------------------------
        // Optional: print to a specific printer with a custom page range.
        // ------------------------------------------------------------
        //PrinterSettings settings = new PrinterSettings
        //{
        //    PrinterName = "My Printer",
        //    PrintRange = PrintRange.SomePages,
        //    FromPage = 1,
        //    ToPage = 2
        //};
        //doc.Print(settings);
    }
}
