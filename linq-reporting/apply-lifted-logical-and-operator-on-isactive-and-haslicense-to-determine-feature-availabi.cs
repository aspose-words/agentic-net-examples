using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class FeatureModel
{
    // Nullable booleans to demonstrate lifted logical AND.
    public bool? IsActive { get; set; } = false;
    public bool? HasLicense { get; set; } = false;

    // Lifted logical AND – result is null if either operand is null.
    // The '&' operator is the lifted version for nullable booleans.
    public bool? FeatureAvailable => IsActive & HasLicense;
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with a LINQ Reporting tag.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Feature available: <<[model.FeatureAvailable]>>");
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // 2. Load the template for reporting.
        var doc = new Document(templatePath);

        // 3. Prepare sample data.
        var model = new FeatureModel
        {
            IsActive = true,      // change to false or null to test other outcomes
            HasLicense = true
        };

        // 4. Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
