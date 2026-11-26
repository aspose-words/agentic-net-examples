using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data model.
        var model = new ReportModel { Name = "John Doe" };

        // Create a template document with a LINQ Reporting tag.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello, <<[model.Name]>>!");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report using the model as the root object.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Public data model class required by the template.
public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}
