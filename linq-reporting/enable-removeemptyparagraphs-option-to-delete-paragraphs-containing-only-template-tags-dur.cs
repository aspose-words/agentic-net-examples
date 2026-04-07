using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with a paragraph that contains only a tag.
        const string templateFile = "Template.docx";
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // This paragraph will become empty after the tag is evaluated (Empty property is empty).
        builder.Writeln("<<[model.Empty]>>");

        // A normal paragraph to show that other content remains.
        builder.Writeln("Hello <<[model.Name]>>");

        // Save the template to disk.
        template.Save(templateFile);

        // Load the template for reporting.
        Document reportDoc = new Document(templateFile);

        // Prepare the data model. The Empty property is an empty string, so its tag resolves to nothing.
        ReportModel model = new ReportModel
        {
            Name = "World",
            Empty = string.Empty
        };

        // Configure the reporting engine to remove empty paragraphs.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}

// Simple data model with public properties.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
    public string Empty { get; set; } = string.Empty;
}
