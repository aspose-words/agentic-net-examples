using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TemplateProcessor
{
    /// <summary>
    /// Builds a report from the specified template while disabling reflection optimization.
    /// The original optimization setting is restored after the operation.
    /// </summary>
    /// <param name="templatePath">Path to the template document.</param>
    /// <param name="outputPath">Path where the generated report will be saved.</param>
    /// <param name="dataSource">Data source used by the reporting engine.</param>
    public void BuildReportWithoutReflectionOptimization(string templatePath, string outputPath, object dataSource)
    {
        // Load the template document.
        Document template = new Document(templatePath);

        // Preserve the current global setting.
        bool originalOptimization = ReportingEngine.UseReflectionOptimization;

        // Disable the reflection optimization for this scoped operation.
        ReportingEngine.UseReflectionOptimization = false;

        try
        {
            // Create a new reporting engine instance.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the provided data source.
            engine.BuildReport(template, dataSource);

            // Save the generated report.
            template.Save(outputPath);
        }
        finally
        {
            // Restore the original optimization setting regardless of success or failure.
            ReportingEngine.UseReflectionOptimization = originalOptimization;
        }
    }
}

public static class Program
{
    /// <summary>
    /// Entry point required for a console application. Demonstrates the scoped disabling of reflection optimization.
    /// </summary>
    public static void Main(string[] args)
    {
        // Example file paths – replace with actual locations when running the sample.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Simple data source; in real scenarios this could be a DataSet, List<T>, etc.
        var dataSource = new
        {
            Title = "Quarterly Report",
            Date = DateTime.Now,
            Items = new[]
            {
                new { Name = "Item A", Quantity = 10 },
                new { Name = "Item B", Quantity = 20 }
            }
        };

        var processor = new TemplateProcessor();
        processor.BuildReportWithoutReflectionOptimization(templatePath, outputPath, dataSource);

        Console.WriteLine($"Report generated and saved to '{outputPath}'.");
    }
}
