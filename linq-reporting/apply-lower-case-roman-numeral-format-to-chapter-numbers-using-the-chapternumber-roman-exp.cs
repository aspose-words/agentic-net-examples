using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Chapter
{
    public int Number { get; set; }
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
        // 1. Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Chapters collection.
        builder.Writeln("<<foreach [chapter in Chapters]>>");
        // Use the roman format (lower‑case) for the chapter number.
        builder.Writeln("Chapter <<[chapter.Number]:roman>>: <<[chapter.Title]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template for reporting.
        Document doc = new Document(templatePath);

        // 3. Prepare sample data.
        ReportModel model = new ReportModel
        {
            Chapters = new List<Chapter>
            {
                new Chapter { Number = 1, Title = "Introduction" },
                new Chapter { Number = 2, Title = "Getting Started" },
                new Chapter { Number = 3, Title = "Advanced Topics" }
            }
        };

        // 4. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        doc.Save("Report.docx");
    }
}
