using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create template document with LINQ Reporting tags.
        var templatePath = "template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Report of Sections");
        builder.Writeln();
        builder.Writeln("<<foreach [sec in Sections]>>");
        builder.Writeln("<<[sec.SectionLetter]>>. <<[sec.Title]>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Prepare data model.
        var model = new ReportModel
        {
            Sections = new List<Section>
            {
                new Section { SectionNumber = 1, Title = "Introduction" },
                new Section { SectionNumber = 2, Title = "Methodology" },
                new Section { SectionNumber = 3, Title = "Results" },
                new Section { SectionNumber = 4, Title = "Conclusion" }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "report.docx";
        doc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

// Section data model.
public class Section
{
    public int SectionNumber { get; set; }
    public string Title { get; set; } = string.Empty;

    // Convert integer to uppercase alphabetic letter (A, B, C, ...).
    public string SectionLetter => ((char)('A' + SectionNumber - 1)).ToString();
}
