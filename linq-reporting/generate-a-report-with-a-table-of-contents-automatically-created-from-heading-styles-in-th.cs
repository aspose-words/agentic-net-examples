using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingToc
{
    // Simple wrapper model required by the ReportingEngine.
    public class ReportModel
    {
        // Example collection; not used in this simple scenario but demonstrates a realistic model.
        public List<SectionItem> Sections { get; set; } = new();
    }

    public class SectionItem
    {
        public string Title { get; set; } = string.Empty;
        public int Level { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the final report.
            const string templatePath = "Template.docx";
            const string reportPath = "ReportWithTOC.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a Table of Contents that will pick up headings 1‑3.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // Add sample headings with built‑in heading styles.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 1 – Introduction");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.1 – Background");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 1.2 – Scope");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Chapter 2 – Implementation");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Section 2.1 – Architecture");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            builder.Writeln("Subsection 2.1.1 – Components");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Create a dummy data model (no LINQ tags are required for this example).
            ReportModel model = new ReportModel();

            // Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The root object name must match the name used in the template
            // (the template does not reference the model, so the name can be arbitrary).
            bool success = engine.BuildReport(reportDoc, model, "model");

            // Update fields so the TOC reflects the generated headings.
            if (success)
            {
                reportDoc.UpdateFields();
            }

            // Save the final document.
            reportDoc.Save(reportPath);
        }
    }
}
