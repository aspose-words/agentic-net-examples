// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Drawing.Printing;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Rendering;

class Program
{
    static void Main()
    {
        // 1. Load a Word template that contains Aspose.Words Reporting tags.
        //    The Document(string) constructor is the approved way to create a document from a file.
        Document doc = new Document("Template.docx");

        // 2. Prepare a simple data source for the report.
        //    Any .NET object (including anonymous types) can be used with ReportingEngine.
        var dataSource = new
        {
            Title = "Sales Report",
            GeneratedOn = DateTime.Now,
            Items = new[]
            {
                new { Product = "Apple",  Quantity = 10, Price = 1.20 },
                new { Product = "Banana", Quantity = 5,  Price = 0.80 },
                new { Product = "Orange", Quantity = 8,  Price = 1.10 }
            }
        };

        // 3. Populate the template with the data.
        //    ReportingEngine.BuildReport(Document, object) follows the provided rule.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource);

        // 4. Save the populated document as MHTML.
        //    Use the Save(string, SaveFormat) overload with SaveFormat.Mhtml.
        doc.Save("Report.mhtml", SaveFormat.Mhtml);

        // 5. Print the document programmatically.
        //    AsposeWordsPrintDocument wraps the Document for .NET printing.
        AsposeWordsPrintDocument printDoc = new AsposeWordsPrintDocument(doc);

        // Optional: specify a printer name if you do not want to use the default printer.
        // printDoc.PrinterSettings.PrinterName = "Your Printer Name";

        // Print using the standard (no UI) print controller.
        printDoc.Print();
    }
}
