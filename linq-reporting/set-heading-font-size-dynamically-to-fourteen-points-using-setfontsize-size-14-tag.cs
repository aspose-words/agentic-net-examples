using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Title { get; set; } = "Dynamic Heading";
}

public class Program
{
    public static void Main()
    {
        // Create sample data.
        var model = new ReportModel();

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Set the font size for the heading to 14 points.
        builder.Font.Size = 14;

        // Insert a heading placeholder that will be replaced by the model's Title.
        builder.Writeln("<<[model.Title]>>");

        // Save the template to a local file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
