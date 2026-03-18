using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Disable the reflection optimization. This is useful when working with small data sets,
        // as the overhead of dynamic class generation can outweigh the benefit of direct reflection.
        ReportingEngine.UseReflectionOptimization = false;

        // Load a template document. Replace the path with the actual location of your template.
        Document template = new Document("Template.docx");

        // Create a simple data source. In real scenarios this could be any object, collection, or custom data source.
        var data = new
        {
            Name = "John Doe",
            Address = "123 Main St, Anytown"
        };

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, data);

        // Save the generated report.
        template.Save("Report.docx");
    }
}
