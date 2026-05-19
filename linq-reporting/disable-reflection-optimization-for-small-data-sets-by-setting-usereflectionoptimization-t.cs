using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare a simple data model
        var model = new ReportModel { Name = "John Doe" };

        // Create a template document containing a LINQ Reporting tag
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, <<[model.Name]>>!");

        // Disable reflection optimization for small data sets
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report using the template and the data model
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save("ReportOutput.docx");
    }

    // Simple public data model with a non‑nullable property
    public class ReportModel
    {
        public string Name { get; set; } = "";
    }
}
