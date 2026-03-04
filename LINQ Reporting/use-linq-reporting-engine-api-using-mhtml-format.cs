using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the Word template that contains LINQ Reporting Engine tags.
        Document doc = new Document("Template.docx");

        // Prepare a simple anonymous data source.
        var data = new
        {
            Title = "Sales Report",
            Date = DateTime.Now,
            Items = new[]
            {
                new { Product = "Apple",  Quantity = 10, Price = 1.20 },
                new { Product = "Banana", Quantity = 5,  Price = 0.80 }
            }
        };

        // Build the report by populating the template with the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "ds"); // "ds" is the name used in the template.

        // Save the populated document as MHTML (Web archive) using HtmlSaveOptions.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        doc.Save("Report.mhtml", saveOptions);
    }
}
