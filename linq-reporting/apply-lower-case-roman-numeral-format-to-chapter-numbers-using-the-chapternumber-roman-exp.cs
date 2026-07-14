using System;
using System.Collections.Generic;
using System.IO;
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
        // Prepare sample data
        var model = new ReportModel
        {
            Chapters = new List<Chapter>
            {
                new Chapter { ChapterNumber = 1, Title = "Introduction" },
                new Chapter { ChapterNumber = 2, Title = "Getting Started" },
                new Chapter { ChapterNumber = 3, Title = "Advanced Topics" }
            }
        };

        // Create template document
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Book Chapters:");
        builder.Writeln("<<foreach [ch in Chapters]>>");
        // Apply lower‑case Roman numeral format to ChapterNumber
        builder.Writeln("{=ch.ChapterNumber:roman}. <<[ch.Title]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load template and build report
        var templateDoc = new Document(templatePath);
        ReportingEngine.UseReflectionOptimization = true;
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        var outputPath = "Report.docx";
        templateDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
