// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document from the file system.
        // This uses the Document(string) constructor – the standard load rule.
        Document doc = new Document("Input.docx");

        // Ensure that any fields (e.g., DATE, PAGE) are up‑to‑date before printing.
        doc.UpdateFields();

        // Print the whole document to the default printer.
        // The Print() method is part of the Document API.
        doc.Print();

        // ------------------------------------------------------------
        // Optional: print to a specific printer with custom settings.
        // Uncomment and adjust the printer name as needed.
        // ------------------------------------------------------------
        // PrinterSettings printerSettings = new PrinterSettings();
        // printerSettings.PrinterName = "MyPrinter";
        // printerSettings.PrintRange = PrintRange.AllPages;
        // doc.Print(printerSettings, "PrintedDocument");
    }
}
