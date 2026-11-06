using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingRomanNumerals
{
    public class Chapter
    {
        public int ChapterNumber { get; set; } = 0;
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
            // Register code page provider for any legacy encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new ReportModel
            {
                Chapters = new List<Chapter>
                {
                    new Chapter { ChapterNumber = 1, Title = "Introduction" },
                    new Chapter { ChapterNumber = 2, Title = "Getting Started" },
                    new Chapter { ChapterNumber = 3, Title = "Advanced Topics" }
                }
            };

            // Create a template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Book Chapters");
            builder.Writeln();

            // LINQ Reporting tags.
            builder.Writeln("<<foreach [c in Chapters]>>");
            // Apply lower‑case Roman numeral format to the chapter number.
            builder.Writeln("{=c.ChapterNumber:roman}. <<[c.Title]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(templatePath);

            // Load the template for reporting.
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root object named "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
