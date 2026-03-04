// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Rendering;
using System.Drawing.Printing;

class Program
{
    static void Main()
    {
        // Load the DOTM template from disk.
        Document doc = new Document("Template.dotm");

        // Rebuild the page layout so that pagination is current before printing.
        doc.UpdatePageLayout();

        // Print the whole document to the default printer.
        doc.Print();

        // -------------------------------------------------
        // If you need to print to a specific printer or
        // control printing options, use AsposeWordsPrintDocument.
        // -------------------------------------------------
        //PrinterSettings printerSettings = new PrinterSettings();
        //printerSettings.PrinterName = "Your Printer Name";
        //AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
        //awPrintDoc.PrinterSettings = printerSettings;
        //awPrintDoc.Print();
    }
}
