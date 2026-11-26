using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple data model with only one existing member.
        var model = new SampleModel
        {
            ExistingMember = "Hello World"
        };

        // Create the template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Write a tag that references an existing member.
        builder.Writeln("Existing: <<[model.ExistingMember]>>");
        // Write a tag that references a missing member.
        builder.Writeln("Missing: <<[model.MissingMember]>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Configure the reporting engine to allow missing members.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        // Verify that the missing member was rendered as a null literal.
        string reportText = reportDoc.GetText();
        bool containsNullLiteral = reportText.Contains("null", StringComparison.OrdinalIgnoreCase);

        Console.WriteLine($"Missing member rendered as null literal: {containsNullLiteral}");
    }
}

// Public data model class.
public class SampleModel
{
    // Existing member initialized to avoid nullable warnings.
    public string ExistingMember { get; set; } = string.Empty;
}
