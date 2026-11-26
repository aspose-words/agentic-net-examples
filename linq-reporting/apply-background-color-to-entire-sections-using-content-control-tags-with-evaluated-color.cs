using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingBackgroundColor
{
    // Model classes
    public class ReportModel
    {
        public List<Section> Sections { get; set; } = new();
    }

    public class Section
    {
        public string Title { get; set; } = "";
        public string Content { get; set; } = "";
        public string BgColor { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data
            var model = new ReportModel
            {
                Sections = new List<Section>
                {
                    new Section
                    {
                        Title = "Introduction",
                        Content = "This is the introduction section.",
                        BgColor = "LightYellow"
                    },
                    new Section
                    {
                        Title = "Conclusion",
                        Content = "This is the conclusion section.",
                        BgColor = "#ADD8E6" // LightBlue
                    }
                }
            };

            // Create template document
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Insert LINQ Reporting tags
            builder.Writeln("<<foreach [sec in Sections]>>");
            builder.Writeln("<<backColor [sec.BgColor]>>");
            builder.Writeln("<<[sec.Title]>>");
            builder.Writeln("<br/>");
            builder.Writeln("<<[sec.Content]>>");
            builder.Writeln("<</backColor>>");
            builder.Writeln("<</foreach>>");

            // Save template (optional, can be omitted)
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load template for reporting (demonstrates load step)
            var doc = new Document(templatePath);

            // Build report
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save final report
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
