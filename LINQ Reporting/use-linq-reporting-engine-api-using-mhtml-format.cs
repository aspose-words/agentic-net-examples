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
            Items = new[]
            {
                new { Name = "Product A", Quantity = 120 },
                new { Name = "Product B", Quantity = 85 },
                new { Name = "Product C", Quantity = 47 }
            }
        };

        // Populate the template with the data source.
        ReportingEngine engine = new ReportingEngine();
        // The third argument is the name used to reference the data source inside the template (e.g. <<[ds.Title]>>).
        engine.BuildReport(doc, data, "ds");

        // Save the populated document as MHTML (Web archive) using HtmlSaveOptions.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml);
        doc.Save("Report.mhtml", saveOptions);
    }
}
