using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        ReportModel model = new()
        {
            Chapters = new()
            {
                new Chapter { Number = 1, Title = "Introduction" },
                new Chapter { Number = 2, Title = "Getting Started" },
                new Chapter { Number = 3, Title = "Advanced Topics" }
            }
        };

        // Create the template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report of Chapters");
        builder.Writeln();
        builder.Writeln("<<foreach [c in Chapters]>>");
        builder.Writeln("Chapter <<[c.Number]:roman>>: <<[c.Title]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template and build the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

public class ReportModel
{
    public List<Chapter> Chapters { get; set; } = new();
}

public class Chapter
{
    public int Number { get; set; }
    public string Title { get; set; } = string.Empty;
}
