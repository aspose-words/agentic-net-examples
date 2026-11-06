using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create template document
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // Load template
        Document doc = new Document(templatePath);

        // Prepare sample data
        ReportModel model = new()
        {
            Sections = new()
            {
                new Section { Title = "Introduction", Content = "This is the introduction section." },
                new Section { Title = "Chapter 1", Content = "Details of chapter 1 go here." },
                new Section { Title = "Conclusion", Content = "Final thoughts are presented in this section." }
            }
        };

        // Build report using LINQ Reporting Engine
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Update fields to generate the Table of Contents
        doc.UpdateFields();

        // Save the final report
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Insert Table of Contents field
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.Writeln(); // Add a blank line after TOC

        // Begin foreach over Sections
        builder.Writeln("<<foreach [section in Sections]>>");

        // Heading for each section (Heading 1 style)
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[section.Title]>>");

        // Normal paragraph for section content
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("<<[section.Content]>>");

        // End foreach
        builder.Writeln("<</foreach>>");

        // Save template
        doc.Save(path);
    }
}

// Data model for the report
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = string.Empty;
    public string Content { get; set; } = string.Empty;
}
