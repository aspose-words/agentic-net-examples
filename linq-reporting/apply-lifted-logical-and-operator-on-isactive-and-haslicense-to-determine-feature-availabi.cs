using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class FeatureModel
{
    // Nullable booleans to demonstrate lifted logical operations.
    public bool? IsActive { get; set; } = true;
    public bool? HasLicense { get; set; } = true;
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";

        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Feature Availability Report");

        // Use null‑coalescing to safely evaluate nullable booleans.
        builder.Writeln("<<if [(model.IsActive ?? false) && (model.HasLicense ?? false)]>>Feature is AVAILABLE<</if>>");
        builder.Writeln("<<if [!((model.IsActive ?? false) && (model.HasLicense ?? false))]>>Feature is NOT AVAILABLE<</if>>");

        // Save the template so it can be loaded for reporting.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample data.
        // -----------------------------------------------------------------
        var model = new FeatureModel
        {
            IsActive = true,
            HasLicense = false   // Change values to see different outcomes.
        };

        // -----------------------------------------------------------------
        // 4. Build the report.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
