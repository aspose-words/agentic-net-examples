using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace RestartNumberingExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Section> Sections { get; set; } = new();
    }

    public class Section
    {
        public string Title { get; set; } = string.Empty;
        public List<string> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data
            var model = new ReportModel
            {
                Sections = new List<Section>
                {
                    new Section
                    {
                        Title = "Section A",
                        Items = new List<string> { "Item A1", "Item A2", "Item A3" }
                    },
                    new Section
                    {
                        Title = "Section B",
                        Items = new List<string> { "Item B1", "Item B2", "Item B3", "Item B4" }
                    }
                }
            };

            // Create template document programmatically
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Begin outer foreach over sections
            builder.Writeln("<<foreach [section in Model.Sections]>>");
            // Section title
            builder.Writeln("<<[section.Title]>>");
            // Numbered list with restartNum before inner foreach
            builder.Writeln("1. <<restartNum>><<foreach [item in section.Items]>> <<[item]>>");
            // Close inner foreach
            builder.Writeln("<</foreach>>");
            // Close outer foreach
            builder.Writeln("<</foreach>>");

            // Save the template
            doc.Save(templatePath);

            // Load the template for reporting
            var template = new Document(templatePath);

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(template, model, "Model");

            // Save the generated report
            var outputPath = "Report.docx";
            template.Save(outputPath);
        }
    }
}
