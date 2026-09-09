using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class TagReportModel
{
    // Collection of tags to be displayed.
    public List<string> Tags { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare a simple template document programmatically.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Prepare the data model with some tags.
        TagReportModel model = new TagReportModel
        {
            Tags = new List<string> { "aspnet", "csharp", "linq", "reporting" }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Tag List Report");
        builder.Writeln();

        // Conditional block: display tags only if the collection is not empty.
        builder.Writeln("<<if [model.Tags.Any()]>>");
        builder.Writeln("Tags:");
        builder.Writeln("<<foreach [tag in model.Tags]>>");
        builder.Writeln("- <<[tag]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</if>>");

        // Save the template.
        doc.Save(filePath);
    }
}
