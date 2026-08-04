using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            const string templatePath = "Template.docx";

            // Create a blank document and a builder to insert content.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over the Chapters collection.
            builder.Writeln("<<foreach [chapter in Chapters]>>");

            // Use the built‑in "roman" format to display the chapter number in lower‑case Roman numerals.
            // Correct syntax: <<[expression]:format>>
            builder.Writeln("Chapter <<[chapter.ChapterNumber]:roman>>: <<[chapter.Title]>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document that will be populated.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare sample data for the report.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Chapters = new List<Chapter>
                {
                    new Chapter { ChapterNumber = 1, Title = "Introduction" },
                    new Chapter { ChapterNumber = 2, Title = "Getting Started" },
                    new Chapter { ChapterNumber = 3, Title = "Advanced Topics" }
                }
            };

            // -----------------------------------------------------------------
            // 4. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }

    // Root data model exposed to the template.
    public class ReportModel
    {
        public List<Chapter> Chapters { get; set; } = new();
    }

    // Simple chapter class.
    public class Chapter
    {
        public int ChapterNumber { get; set; }
        public string Title { get; set; } = string.Empty;
    }
}
