using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Paths to the template and the output document.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Load the template document. Document does not implement IDisposable in older
        // Aspose.Words versions, so we avoid a using statement.
        Document doc = new Document(templatePath);
        try
        {
            // Disable the reflection optimization for this specific report generation.
            ReportingEngine.UseReflectionOptimization = false;

            // Create a reporting engine instance.
            ReportingEngine engine = new ReportingEngine();

            // Example data source – replace with your actual data.
            var data = new
            {
                Title = "Sample Report",
                Date = DateTime.Now,
                Items = new[]
                {
                    new { Name = "Item 1", Quantity = 10 },
                    new { Name = "Item 2", Quantity = 20 }
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
