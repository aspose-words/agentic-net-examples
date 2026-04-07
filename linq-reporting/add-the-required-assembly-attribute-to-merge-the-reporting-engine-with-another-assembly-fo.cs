using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Name = "John Doe",
            Age = 30
        };

        // Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[model.Age]>>");

        // Save the template (demonstrates the lifecycle rule).
        template.Save("template.docx");

        // Load the template again for report generation.
        var doc = new Document("template.docx");

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("report.docx");
    }
}

// Public data model with initialized properties to avoid nullable warnings.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
