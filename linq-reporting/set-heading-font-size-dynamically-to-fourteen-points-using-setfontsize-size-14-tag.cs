using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Title that will be inserted into the report.
    public string Title { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Set the desired font size for the heading.
        builder.Font.Size = 14;

        // Insert a placeholder that will be replaced by the model's Title.
        builder.Writeln("<<[model.Title]>>");

        // Save the template to disk (required before building the report).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Create a model instance with sample data.
        ReportModel model = new ReportModel
        {
            Title = "Dynamic Heading – Fourteen Points"
        };

        // -----------------------------------------------------------------
        // 3. Build the report using Aspose.Words LINQ Reporting Engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name must match the tag reference (model).
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
