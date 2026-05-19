using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Restrict types that provide file‑writing capabilities.
        // This must be done before any report is built.
        ReportingEngine.SetRestrictedTypes(typeof(System.IO.File), typeof(System.IO.Directory));

        // -----------------------------------------------------------------
        // Create a simple template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting tag that references the data model.
        builder.Writeln("Report for <<[model.Name]>>");

        // Save the template to disk (required before loading for BuildReport).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back into a new Document instance.
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Prepare the data source.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel { Name = "Alice" };

        // -----------------------------------------------------------------
        // Build the report.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            // Allow missing members so that any restricted calls are ignored
            // rather than throwing an exception.
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Populate the document with data from the model.
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
