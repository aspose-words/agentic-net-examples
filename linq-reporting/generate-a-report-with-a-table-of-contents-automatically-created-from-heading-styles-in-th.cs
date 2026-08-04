using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a template document with a TOC and some headings.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a Table of Contents that will pick up headings 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample headings. The TOC will be generated from these.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Overview");
        builder.Writeln("Section 1.2: Details");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Usage");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1: Installation");
        builder.Writeln("Section 2.2: Configuration");

        // Save the template to disk.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and run the LINQ Reporting engine.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);

        // The template does not contain any LINQ Reporting tags, but we still
        // invoke the engine to follow the required workflow.
        ReportingEngine engine = new ReportingEngine();

        // Use a simple wrapper object as the data source.
        ReportModel model = new();
        engine.BuildReport(report, model, "model");

        // -----------------------------------------------------------------
        // Step 3: Update fields so the TOC reflects the headings.
        // -----------------------------------------------------------------
        report.UpdateFields();

        // Save the final report.
        report.Save(reportPath);
    }
}

// Simple wrapper class required by the ReportingEngine call.
public class ReportModel
{
    public List<Chapter> Chapters { get; set; } = new();
}

// Represents a chapter/section – not used directly in this example
// but demonstrates a realistic data model for LINQ Reporting scenarios.
public class Chapter
{
    public string Title { get; set; } = string.Empty;
    public int Level { get; set; }
}
