using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCM template that contains reporting tags.
        Document template = new Document("Template.docm");

        // Example data source – can be any non‑dynamic .NET type.
        var dataSource = new
        {
            Title = "Quarterly Report",
            ReportDate = DateTime.Now,
            Items = new[]
            {
                new { Name = "Product A", Quantity = 120, Price = 9.99 },
                new { Name = "Product B", Quantity = 85,  Price = 14.50 }
            }
        };

        // Populate the template using the ReportingEngine.
        // The third argument ("ds") is the name used to reference the data source in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "ds");

        // Save the generated report.
        template.Save("Result.docx");
    }
}
