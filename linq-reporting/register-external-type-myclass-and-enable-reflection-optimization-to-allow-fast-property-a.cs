using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class MyClass
{
    // Initialize properties to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization for faster property access.
        ReportingEngine.UseReflectionOptimization = true;

        // Create a template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>! Value: <<[model.Value]>>.");

        // Save the template.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Prepare the data model.
        MyClass model = new MyClass
        {
            Name = "Aspose",
            Value = 123
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
    }
}
