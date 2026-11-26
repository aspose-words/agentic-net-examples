using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

namespace LinqReportingTocExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Section> Sections { get; set; } = new();
    }

    public class Section
    {
        public string Title { get; set; } = "";
        public string Content { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for template and output
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // Step 1: Create the template document with a TOC field and placeholders
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a TOC field at the beginning of the document (LINQ Reporting does not have a <<toc>> tag)
            builder.InsertField("TOC \\o \"1-3\" \\h \\z \\u", "Table of Contents");
            builder.Writeln(); // Ensure a blank line after the TOC

            // Begin foreach loop over Sections collection
            builder.Writeln("<<foreach [sec in Sections]>>");

            // Insert a heading for each section (Heading 1 style is recognized by the TOC)
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("<<[sec.Title]>>");

            // Insert normal paragraph with section content
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("<<[sec.Content]>>");
            builder.InsertParagraph(); // Add spacing between sections

            // End foreach loop
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report using LINQ Reporting
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // Prepare sample data
            var model = new ReportModel
            {
                Sections = new List<Section>
                {
                    new Section
                    {
                        Title = "Introduction",
                        Content = "This is the introduction section of the report."
                    },
                    new Section
                    {
                        Title = "Analysis",
                        Content = "Detailed analysis goes here."
                    },
                    new Section
                    {
                        Title = "Conclusion",
                        Content = "Final thoughts and conclusions."
                    }
                }
            };

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Update fields (TOC) after the report is built
            doc.UpdateFields();

            // Save the final document
            doc.Save(outputPath);
        }
    }
}
