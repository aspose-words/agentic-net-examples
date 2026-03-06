using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOC template that contains LINQ Reporting Engine tags.
        Document template = new Document("Template.docx");

        // Prepare a simple data source. The object can be any non‑dynamic, non‑anonymous type.
        var reportData = new ReportData
        {
            Title = "Quarterly Sales Report",
            Items = new[]
            {
                new ReportItem { Name = "Product A", Quantity = 120, Price = 9.99 },
                new ReportItem { Name = "Product B", Quantity = 85,  Price = 14.50 },
                new ReportItem { Name = "Product C", Quantity = 60,  Price = 7.25 }
            }
        };

        // Create the reporting engine, optionally set options (e.g., remove empty paragraphs).
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Populate the template with the data source. The third argument is the name used in the template.
        engine.BuildReport(template, reportData, "model");

        // Save the generated report in DOCX format.
        template.Save("Report.docx", SaveFormat.Docx);
    }
}

// Simple POCO classes used as the data source.
public class ReportData
{
    public string Title { get; set; }
    public ReportItem[] Items { get; set; }
}

public class ReportItem
{
    public string Name { get; set; }
    public int Quantity { get; set; }
    public double Price { get; set; }
}
