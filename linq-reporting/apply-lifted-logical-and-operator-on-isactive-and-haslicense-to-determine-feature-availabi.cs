using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare the data model.
        var model = new ReportModel
        {
            IsActive = true,
            HasLicense = null // Nullable bool to demonstrate lifted logical AND behavior.
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Feature available (IsActive && HasLicense): <<[model.FeatureAvailable]>>");
        doc.Save(templatePath);

        // Load the template back before building the report.
        var loadedDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(loadedDoc, model, "model");

        // Save the final report.
        var outputPath = "Report.docx";
        loadedDoc.Save(outputPath);
    }
}

// Public data model aligned with the template.
public class ReportModel
{
    // Nullable booleans to allow lifted logical AND.
    public bool? IsActive { get; set; } = false;
    public bool? HasLicense { get; set; } = false;

    // Feature availability using lifted && operator.
    // The single '&' operator works with nullable booleans and returns a nullable result.
    public bool? FeatureAvailable => IsActive & HasLicense;
}
