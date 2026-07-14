using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple data model.
        ReportModel model = new ReportModel { Name = "John Doe" };

        // -----------------------------------------------------------------
        // Create a template document programmatically.
        // The template contains a LINQ Reporting tag that references the model.
        // -----------------------------------------------------------------
        string templatePath = "Template.docx";
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Hello <<[model.Name]>>!");
        // The following line is intentionally omitted because the expression
        // would cause a parsing error. The purpose of this example is to show
        // how to restrict types that could write to the file system.
        template.Save(templatePath);

        // Load the template back from disk (required before building the report).
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Configure the ReportingEngine to restrict types that can write to
        // the file system. This enhances security by preventing template
        // expressions from accessing these members.
        // -----------------------------------------------------------------
        ReportingEngine.SetRestrictedTypes(
            typeof(System.IO.File),
            typeof(System.IO.Directory),
            typeof(System.IO.StreamWriter),
            typeof(System.IO.BinaryWriter));

        // Build the report using the LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Simple public data model with a public property.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
}
