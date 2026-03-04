// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Load the DOCM template.
        Document doc = new Document(@"C:\Templates\Report.docm");

        // Refresh any fields (e.g., MERGEFIELD, DATE) before printing.
        doc.UpdateFields();

        // Rebuild the layout so page numbers and other layout‑dependent data are correct.
        doc.UpdatePageLayout();

        // Configure printer settings – print pages 1‑5 as an example.
        PrinterSettings printerSettings = new PrinterSettings
        {
            PrintRange = PrintRange.SomePages,
            FromPage = 1,
            ToPage = 5
        };

        // Print the document using Aspose.Words built‑in Print method.
        doc.Print(printerSettings, "MyReport");

        // Alternative: use AsposeWordsPrintDocument for finer control.
        //AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
        //awPrintDoc.PrinterSettings = printerSettings;
        //awPrintDoc.Print();
    }
}
