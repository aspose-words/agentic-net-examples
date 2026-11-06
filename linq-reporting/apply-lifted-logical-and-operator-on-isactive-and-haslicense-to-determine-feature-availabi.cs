using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class FeatureModel
{
    // Nullable booleans to demonstrate lifted logical AND.
    public bool? IsActive { get; set; } = false;
    public bool? HasLicense { get; set; } = false;
}

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "LinQReportingExample");
        Directory.CreateDirectory(workDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "report.docx");

        // ---------- Create the template document ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Feature Availability Report");
        builder.Writeln();

        // Conditional block using lifted logical AND on nullable booleans.
        // Use explicit comparison to true because the reporting engine cannot apply && to nullable booleans.
        builder.Writeln("<<if [model.IsActive == true && model.HasLicense == true]>>Feature is AVAILABLE.<</if>>");
        builder.Writeln("<<if [!(model.IsActive == true && model.HasLicense == true)]>>Feature is NOT available.<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Load the template for reporting ----------
        Document doc = new Document(templatePath);

        // Prepare sample data.
        FeatureModel model = new FeatureModel
        {
            IsActive = true,
            HasLicense = true
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
