using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with a LINQ Reporting link tag.
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        // The link tag takes a URI expression and an optional display text expression.
        builder.Writeln("<<link [model.Url] [model.Text]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare the data model.
        var model = new ReportModel
        {
            Url = "https://www.example.com",
            Text = "Visit Example"
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}

// Simple data model with public properties.
public class ReportModel
{
    public string Url { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
}
