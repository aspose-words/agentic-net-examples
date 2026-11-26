using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
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
            string templatePath = "Template.docx";
            string outputPath = "ReportWithToc.docx";

            // 1. Create the template document programmatically
            CreateTemplate(templatePath);

            // 2. Load the template
            Document doc = new Document(templatePath);

            // 3. Prepare sample data
            ReportModel model = new ReportModel
            {
                Sections = new List<Section>
                {
                    new Section { Title = "Introduction", Content = "This is the introduction section." },
                    new Section { Title = "Chapter 1", Content = "Details of chapter 1 go here." },
                    new Section { Title = "Chapter 2", Content = "Details of chapter 2 go here." },
                    new Section { Title = "Conclusion", Content = "Final thoughts and conclusions." }
                }
            };

            // 4. Build the report using LINQ Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 5. Update fields to generate the Table of Contents
            doc.UpdateFields();

            // 6. Save the final report
            doc.Save(outputPath);

            // Inform the user (no interactive input)
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }

        private static void CreateTemplate(string path)
        {
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert Table of Contents field
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.Writeln(); // Add a blank paragraph after TOC

            // Begin foreach over Sections
            builder.Writeln("<<foreach [section in Sections]>>");

            // Heading styled as Heading 1 with LINQ tag for title
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("<<[section.Title]>>");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

            // Content paragraph
            builder.Writeln("<<[section.Content]>>");

            // End foreach
            builder.Writeln("<</foreach>>");

            // Save the template
            templateDoc.Save(path);
        }
    }
}
