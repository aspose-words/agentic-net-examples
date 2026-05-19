using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "FinancialTemplate.docx";
        string reportPath = "FinancialReport.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("Financial Report");
        builder.Writeln();

        // Insert LINQ Reporting tags that use System.Math static members.
        // The ReportingEngine will be configured with Math as a known type,
        // allowing the template to call Math.Round, Math.Floor, etc.
        builder.Writeln("Original Amount: <<[Amount]>>");
        builder.Writeln("Rounded (2 decimals): <<[Math.Round(Amount, 2)]>>");
        builder.Writeln("Rounded (nearest integer): <<[Math.Round(Amount)]>>");
        builder.Writeln("Floor value: <<[Math.Floor(Amount)]>>");
        builder.Writeln("Ceiling value: <<[Math.Ceiling(Amount)]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Sample data model.
        FinancialData data = new FinancialData
        {
            Amount = 1234.5678m
        };

        // Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // Register System.Math so its static members can be used in expressions.
        engine.KnownTypes.Add(typeof(Math));

        // Build the report. The root object name must match the tags (no prefix needed here).
        engine.BuildReport(loadedTemplate, data, "data");

        // Save the generated report.
        loadedTemplate.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class FinancialData
{
    // Initialize to avoid nullable warnings.
    public decimal Amount { get; set; } = 0m;
}
