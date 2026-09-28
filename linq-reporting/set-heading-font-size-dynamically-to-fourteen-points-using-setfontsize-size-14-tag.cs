using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model
        ReportModel model = new()
        {
            HeadingHtml = "<h1 style='font-size:14pt;'>Dynamic Heading</h1>"
        };

        // Create a template document programmatically
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("<<[model.HeadingHtml] -html>>");
        templateDoc.Save(templatePath);

        // Load the template (optional – we can reuse the same document)
        Document doc = new(templatePath);

        // Build the report
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = "output.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model used by the LINQ Reporting engine
public class ReportModel
{
    // HTML snippet that defines a heading with a dynamic font size of 14 points
    public string HeadingHtml { get; set; } = "<h1 style='font-size:14pt;'>Default Heading</h1>";
}
