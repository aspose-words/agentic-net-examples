using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class FeatureModel
{
    public bool IsActive { get; set; }
    public bool HasLicense { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with LINQ Reporting tags.
        // The lifted logical AND operator (&&) works with nullable booleans,
        // but here we use non‑nullable booleans for simplicity.
        builder.Writeln("<<if [model.IsActive && model.HasLicense]>>Feature Available<</if>>");
        builder.Writeln("<<if [! (model.IsActive && model.HasLicense)]>>Feature Unavailable<</if>>");

        // Prepare the data source.
        FeatureModel model = new FeatureModel
        {
            IsActive = true,
            HasLicense = false
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("FeatureReport.docx");
    }
}
