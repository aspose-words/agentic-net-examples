using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

namespace AsposeWordsLinqReportingToc
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of sections that will be rendered.
        public List<SectionItem> Sections { get; set; } = new();
    }

    // Represents a single section with a heading and body text.
    public class SectionItem
    {
        public string Title { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and final report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a Table of Contents field that will pick up headings 1‑3.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // LINQ Reporting tags: iterate over Sections collection.
            builder.Writeln("<<foreach [section in Sections]>>");

            // Heading for each section (uses built‑in Heading 1 style).
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("<<[section.Title]>>");

            // Normal paragraph for the section content.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("<<[section.Content]>>");

            // End of the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data for the report.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Sections = new List<SectionItem>
                {
                    new SectionItem
                    {
                        Title = "Introduction",
                        Content = "This is the introduction section of the report."
                    },
                    new SectionItem
                    {
                        Title = "Analysis",
                        Content = "Here we provide a detailed analysis of the data."
                    },
                    new SectionItem
                    {
                        Title = "Conclusion",
                        Content = "The concluding remarks are presented here."
                    }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report using LINQ Reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after tag processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report; the root object name in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // Update fields so that the TOC reflects the generated headings.
            reportDoc.UpdateFields();

            // -----------------------------------------------------------------
            // 4. Save the final report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
