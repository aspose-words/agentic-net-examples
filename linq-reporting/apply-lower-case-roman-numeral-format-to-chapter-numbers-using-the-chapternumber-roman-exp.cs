using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRomanNumerals
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Sample data model
            var model = new ReportModel
            {
                Chapters = new List<Chapter>
                {
                    new Chapter { ChapterNumber = 1, Title = "Introduction" },
                    new Chapter { ChapterNumber = 2, Title = "Getting Started" },
                    new Chapter { ChapterNumber = 3, Title = "Advanced Topics" }
                }
            };

            // Create the template document
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Book Report");
            builder.Writeln("<<foreach [c in Chapters]>>");
            // Apply lower‑case Roman numeral format using the field code expression
            builder.Writeln("{=c.ChapterNumber:roman}. <<[c.Title]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load the template and build the report
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
            reportDoc.Save(outputPath);
        }
    }

    public class ReportModel
    {
        public List<Chapter> Chapters { get; set; } = new();
    }

    public class Chapter
    {
        public int ChapterNumber { get; set; }
        public string Title { get; set; } = "";
    }
}
