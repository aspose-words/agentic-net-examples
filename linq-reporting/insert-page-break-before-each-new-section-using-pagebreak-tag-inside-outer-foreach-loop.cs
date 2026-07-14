using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model for the report.
    public class ReportModel
    {
        public List<SectionItem> Sections { get; set; } = new();
    }

    public class SectionItem
    {
        public string Title { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Outer foreach iterates over the Sections collection.
            builder.Writeln("<<foreach [section in Sections]>>");
            // Section title.
            builder.Writeln("Section: <<[section.Title]>>");
            // Section content.
            builder.Writeln("Content: <<[section.Content]>>");
            // End of the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data for the report.
            // -----------------------------------------------------------------
            ReportModel model = new()
            {
                Sections = new()
                {
                    new SectionItem { Title = "Introduction", Content = "This is the introduction." },
                    new SectionItem { Title = "Details", Content = "Detailed information goes here." },
                    new SectionItem { Title = "Conclusion", Content = "Final thoughts and summary." }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None // Explicit assignment as required.
            };

            // Build the report using the model; the root name must match the template references.
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
