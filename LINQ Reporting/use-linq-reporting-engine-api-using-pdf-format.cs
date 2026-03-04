using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the Word template that contains LINQ Reporting Engine tags.
        Document template = new Document("Template.docx");

        // Prepare a simple data source that matches the tags in the template.
        var data = new
        {
            Title = "Quarterly Report",
            Date = DateTime.Now,
            Items = new[]
            {
                new { Name = "Product A", Quantity = 120, Price = 9.99 },
                new { Name = "Product B", Quantity = 85,  Price = 14.50 }
            }
        };

        // Populate the template with the data using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, data, "report");

        // Save the resulting document as a PDF file.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        template.Save("Report.pdf", pdfOptions);
    }
}
