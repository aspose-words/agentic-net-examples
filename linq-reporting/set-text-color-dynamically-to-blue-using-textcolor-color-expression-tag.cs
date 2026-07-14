using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Step 1: Create a template document with a textColor tag.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // The tag will use the model's Color property to set the text color.
        builder.Writeln("<<textColor [model.Color]>>Status Text<</textColor>>");
        templateDoc.Save(templatePath);

        // Step 2: Load the template document.
        Document doc = new Document(templatePath);

        // Step 3: Prepare the data model.
        ReportModel model = new ReportModel
        {
            Color = "Blue" // Known color name; the engine will apply this to the text.
        };

        // Step 4: Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Step 5: Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }
}

// Simple data model aligned with the template.
public class ReportModel
{
    // The color expression used in the template.
    public string Color { get; set; } = "Blue";
}
