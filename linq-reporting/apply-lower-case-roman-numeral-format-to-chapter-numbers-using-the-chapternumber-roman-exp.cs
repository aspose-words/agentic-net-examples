using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Chapter
{
    public int ChapterNumber { get; set; }
    public string Title { get; set; } = "";
}

public class ReportModel
{
    public List<Chapter> Chapters { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Required for some encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create the template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Chapters collection.
        builder.Writeln("<<foreach [chapter in Chapters]>>");
        // Apply lower‑case Roman numeral format to ChapterNumber.
        builder.Writeln("Chapter <<[chapter.ChapterNumber]:roman>>: <<[chapter.Title]>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save and reload the template to satisfy lifecycle rules.
        const string templatePath = "Template.docx";
        template.Save(templatePath);
        var doc = new Document(templatePath);

        // Sample data.
        var model = new ReportModel
        {
            Chapters = new List<Chapter>
            {
                new Chapter { ChapterNumber = 1, Title = "Introduction" },
                new Chapter { ChapterNumber = 2, Title = "Getting Started" },
                new Chapter { ChapterNumber = 3, Title = "Advanced Topics" }
            }
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
