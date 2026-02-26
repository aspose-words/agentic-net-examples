// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // Path to the MHTML (Web archive) file that contains the report.
        string mhtmlPath = @"C:\Docs\Report.mht";

        // Load the MHTML document. The Document constructor automatically detects the format.
        Document doc = new Document(mhtmlPath);

        // Update any fields (e.g., MERGEFIELD, DATE) before printing.
        doc.UpdateFields();

        // Configure printer settings – print all pages using the default printer.
        PrinterSettings printerSettings = new PrinterSettings
        {
            PrintRange = PrintRange.AllPages
        };

        // Create an Aspose.Words print document which integrates with .NET printing.
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);
        printDoc.PrinterSettings = printerSettings;

        // Send the document to the printer.
        printDoc.Print();
    }
}
