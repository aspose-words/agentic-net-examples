using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model.
        ReportModel model = new ReportModel
        {
            // Change this list to test both scenarios (empty vs non‑empty).
            Tags = new List<string> { "alpha", "beta", "gamma" }
        };

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // If the Tags collection has any items, display a heading and the list.
        builder.Writeln("<<if [model.Tags.Any()]>>");
        builder.Writeln("Tag list:");
        builder.Writeln("<<foreach [tag in model.Tags]>>- <<[tag]>> <</foreach>>");
        builder.Writeln("<</if>>");

        // Save the template to a temporary file.
        string templatePath = Path.Combine(Path.GetTempPath(), "TagTemplate.docx");
        template.Save(templatePath);

        // Load the template back (simulating a real‑world scenario where the template is stored on disk).
        Document doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Path.GetTempPath(), "TagReport.docx");
        doc.Save(outputPath);

        // Inform the user where the files are located (no interactive input required).
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report saved to:   {outputPath}");
    }
}

// Public data model used by the template.
public class ReportModel
{
    // Initialise the collection to avoid nullable warnings.
    public List<string> Tags { get; set; } = new();
}
