using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Lists;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingToc
{
    // Model that holds the data for the report.
    public class ReportModel
    {
        public string Title { get; set; } = "Document Title";
        public string Section1 { get; set; } = "First Section";
        public string Section2 { get; set; } = "Second Section";
        public string SubSection1 { get; set; } = "First Subsection";
        public string Content { get; set; } = "This is some sample content for the report.";
    }

    class Program
    {
        static void Main()
        {
            // Ensure code page provider is registered (required for some data sources).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Insert a Table of Contents field that will pick up headings 1‑3.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // Heading 1 – uses the Title property.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("<<[model.Title]>>");

            // Heading 2 – first section.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("<<[model.Section1]>>");

            // Heading 3 – subsection under first section.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            builder.Writeln("<<[model.SubSection1]>>");

            // Normal paragraph with content.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("<<[model.Content]>>");

            // Another Heading 2 – second section.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("<<[model.Section2]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            var reportDoc = new Document(templatePath);

            // Populate the model with sample data.
            var model = new ReportModel
            {
                Title = "Automated Table of Contents Example",
                Section1 = "Introduction",
                SubSection1 = "Background",
                Content = "This document demonstrates how to generate a TOC using LINQ Reporting tags.",
                Section2 = "Conclusion"
            };

            // Create the reporting engine and build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Update fields so the TOC reflects the generated headings.
            reportDoc.UpdateFields();

            // Save the final report.
            reportDoc.Save(outputPath);
        }
    }
}
