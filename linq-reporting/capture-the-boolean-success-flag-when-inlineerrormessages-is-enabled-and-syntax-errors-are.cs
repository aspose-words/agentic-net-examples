using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string resultPath = Path.Combine(workDir, "result.docx");

        // -----------------------------------------------------------------
        // 1. Create a template document with a valid tag and an invalid tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Valid expression.
        builder.Writeln("Hello <<[model.Name]>>!");

        // Invalid expression – property 'Missing' does not exist in the model.
        builder.Writeln("This will cause an error: <<[model.Missing]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        Model model = new Model { Name = "World" };

        // -----------------------------------------------------------------
        // 4. Configure ReportingEngine with InlineErrorMessages.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report and capture the success flag.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(resultPath);

        // Output the success flag.
        Console.WriteLine($"Report build success: {success}");
    }
}

// Data model used by the template.
public class Model
{
    public string Name { get; set; } = "";
    // Note: No 'Missing' property – this will trigger an inline error.
}
