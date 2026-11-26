using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace RestartNumberingExample
{
    // Data model for the report
    public class ReportModel
    {
        public List<Section> Sections { get; set; } = new();
    }

    public class Section
    {
        public string Title { get; set; } = string.Empty;
        public List<string> Items { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // Create sample data
            var model = new ReportModel
            {
                Sections = new List<Section>
                {
                    new Section
                    {
                        Title = "First Section",
                        Items = new List<string> { "Item A1", "Item A2", "Item A3" }
                    },
                    new Section
                    {
                        Title = "Second Section",
                        Items = new List<string> { "Item B1", "Item B2", "Item B3" }
                    }
                }
            };

            // Create a template document programmatically
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Begin outer foreach over sections
            builder.Writeln("<<foreach [section in model.Sections]>>");
            // Section title
            builder.Writeln("<<[section.Title]>>");
            // Numbered list of items with restartNum before the inner foreach
            builder.Writeln("1. <<restartNum>><<foreach [item in section.Items]>> <<[item]>> <</foreach>>");
            // End outer foreach
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template for reporting
            var doc = new Document(templatePath);

            // Build the report using LINQ Reporting Engine
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
