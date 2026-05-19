using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data model.
        ReportModel model = new ReportModel
        {
            Name = "John Doe"
        };

        // Create a template document that uses the default LINQ Reporting delimiters (<< >>).
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Build the report. The third argument is the name used in the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = "Report.docx";
        doc.Save(reportPath);
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use the default tag syntax required by Aspose.Words LINQ Reporting Engine.
        builder.Writeln("Hello, <<[model.Name]>>!");

        doc.Save(path);
    }
}

// Simple data model with a non‑nullable property initialized to avoid warnings.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
}
