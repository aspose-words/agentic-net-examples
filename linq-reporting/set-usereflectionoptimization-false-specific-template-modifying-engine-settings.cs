using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportData
{
    public string Title { get; set; }
    public DateTime Date { get; set; }
    public Item[] Items { get; set; }
}

public class Item
{
    public string Name { get; set; }
    public int Quantity { get; set; }
}

class Program
{
    static void Main()
    {
        // Paths to the template and the output document.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Load the template document.
        Document doc = new Document(templatePath);
        try
        {
            // Disable the reflection optimization for this specific report generation.
            ReportingEngine.UseReflectionOptimization = false;

            // Create a reporting engine instance.
            ReportingEngine engine = new ReportingEngine();

            // Example data source – replace with your actual data.
            var data = new ReportData
            {
                Title = "Sample Report",
                Date = DateTime.Now,
                Items = new[]
                {
                    new Item { Name = "Item 1", Quantity = 10 },
                    new Item { Name = "Item 2", Quantity = 20 }
                }
            };

            // Build the report using the template and the data source.
            engine.BuildReport(doc, data);

            // Save the generated report.
            doc.Save(outputPath);
        }
        finally
        {
            // Dispose the document explicitly if the version supports IDisposable.
            (doc as IDisposable)?.Dispose();
        }
    }
}
