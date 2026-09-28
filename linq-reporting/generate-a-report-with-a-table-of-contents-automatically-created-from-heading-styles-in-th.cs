using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model
        ReportModel model = new()
        {
            Sections = new List<Section>
            {
                new()
                {
                    Title = "Introduction",
                    Content = "This is the introduction section.",
                    SubSections = new List<SubSection>
                    {
                        new()
                        {
                            Title = "Background",
                            Content = "Background information goes here."
                        },
                        new()
                        {
                            Title = "Purpose",
                            Content = "Purpose of the document."
                        }
                    }
                },
                new()
                {
                    Title = "Main Content",
                    Content = "Details of the main content.",
                    SubSections = new List<SubSection>
                    {
                        new()
                        {
                            Title = "Topic A",
                            Content = "Discussion of topic A."
                        },
                        new()
                        {
                            Title = "Topic B",
                            Content = "Discussion of topic B."
                        }
                    }
                },
                new()
                {
                    Title = "Conclusion",
                    Content = "Final remarks and summary.",
                    SubSections = new List<SubSection>()
                }
            }
        };

        // Create template document programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert Table of Contents field
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Begin foreach over Sections
        builder.Writeln("<<foreach [sec in Sections]>>");

        // Heading 1 for each section title
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[sec.Title]>>");

        // Normal paragraph for section content
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("<<[sec.Content]>>");

        // Begin foreach over SubSections
        builder.Writeln("<<foreach [sub in sec.SubSections]>>");

        // Heading 2 for each subsection title
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("<<[sub.Title]>>");

        // Normal paragraph for subsection content
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("<<[sub.Content]>>");

        // End inner foreach
        builder.Writeln("<</foreach>>");

        // End outer foreach
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for report generation
        Document report = new Document(templatePath);

        // Build the report using LINQ Reporting Engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Update fields to generate the Table of Contents
        report.UpdateFields();

        // Save the final report
        const string outputPath = "output.docx";
        report.Save(outputPath);
    }
}

// Data model classes
public class ReportModel
{
    public List<Section> Sections { get; set; } = new();
}

public class Section
{
    public string Title { get; set; } = "";
    public string Content { get; set; } = "";
    public List<SubSection> SubSections { get; set; } = new();
}

public class SubSection
{
    public string Title { get; set; } = "";
    public string Content { get; set; } = "";
}
