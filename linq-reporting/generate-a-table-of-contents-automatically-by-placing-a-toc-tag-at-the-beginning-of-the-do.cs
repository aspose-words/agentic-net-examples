using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for template and final report.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the document template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Add some headings that will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is the introduction content.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Detailed information goes here.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1: Subsection");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Subsection content.");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report using LINQ Reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new(templatePath);

        // The model is empty because the template does not reference any data.
        var model = new ReportModel();

        // Optional: enable reflection optimization for better performance.
        ReportingEngine.UseReflectionOptimization = true;

        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Insert a Table of Contents at the beginning of the document.
        DocumentBuilder tocBuilder = new(reportDoc);
        tocBuilder.MoveToDocumentStart();
        tocBuilder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Update fields so the TOC is generated.
        reportDoc.UpdateFields();

        // Save the final report.
        reportDoc.Save(reportPath);
    }

    // Empty wrapper class required by the ReportingEngine.
    public class ReportModel
    {
        // No properties needed for this example.
    }
}
