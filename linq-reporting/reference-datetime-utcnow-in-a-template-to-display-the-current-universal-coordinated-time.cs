using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Property that returns the current UTC time when the model is instantiated.
    public DateTime CurrentUtcNow { get; set; } = DateTime.UtcNow;
}

public class Program
{
    public static void Main()
    {
        // Prepare file paths in the current working directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -----------------------------------------------------------------
        // Step 1: Create the template document with a LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a line that will display the current UTC time using the model.
        builder.Writeln("Current UTC time: <<[model.CurrentUtcNow]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportModel model = new ReportModel();

        ReportingEngine engine = new ReportingEngine();
        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // Step 3: Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
