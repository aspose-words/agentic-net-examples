using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data model.
        var model = new ReportModel
        {
            Sections = new List<Section>
            {
                new Section { Title = "Introduction", Content = "This is the introduction section." },
                new Section { Title = "Chapter 1", Content = "Details of chapter 1 go here." },
                new Section { Title = "Conclusion", Content = "Final thoughts and summary." }
            }
        };

        // Create template document.
        const string templatePath = "Template.docx";
        var builder = new DocumentBuilder();

        // Insert Table of Contents field.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.Writeln(); // Add a blank line after TOC.

        // Begin foreach loop over sections.
        builder.Writeln("<<foreach [sec in Sections]>>");

        // Heading paragraph (Heading 1 style) for each section title.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[sec.Title]>>");

        // Reset style to Normal for content.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("<<[sec.Content]>>");

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        builder.Document.Save(templatePath);

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // Update fields (e.g., TOC) after report generation.
        if (success)
        {
            doc.UpdateFields();
        }

        // Save the final report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = string.Empty;
    public string Content { get; set; } = string.Empty;
}
