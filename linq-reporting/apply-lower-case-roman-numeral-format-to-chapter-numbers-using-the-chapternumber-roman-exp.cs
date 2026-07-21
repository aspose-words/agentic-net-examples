using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Chapter
{
    public int ChapterNumber { get; set; }
    public string Title { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Chapter> Chapters { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add a foreach block that iterates over the Chapters collection.
        builder.Writeln("<<foreach [chapter in Chapters]>>");
        // Use the roman format for the chapter number (lower‑case).
        builder.Writeln("Chapter <<[chapter.ChapterNumber]:roman>>: <<[chapter.Title]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Chapters = new List<Chapter>
            {
                new Chapter { ChapterNumber = 1, Title = "Introduction" },
                new Chapter { ChapterNumber = 2, Title = "Getting Started" },
                new Chapter { ChapterNumber = 3, Title = "Advanced Topics" }
            }
        };

        // Build the report using LINQ Reporting.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
