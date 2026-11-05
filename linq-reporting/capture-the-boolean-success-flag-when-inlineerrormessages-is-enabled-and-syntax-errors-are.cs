using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Sample properties used in the template
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
    // This property does NOT exist in the template on purpose to trigger an error
    // (No need to define it; the missing reference will cause the inline error.)
}

public class Program
{
    public static void Main()
    {
        // File names for the temporary template and the generated report
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -------------------------------------------------
        // 1. Create a template document with a deliberate reference error
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Correct tag – will be processed normally
        builder.Writeln("Hello <<[model.Name]>>!");

        // Intentional error: reference a non‑existent property
        builder.Writeln("This line has an error: <<[model.NonExistentProperty]>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template document for reporting
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data model
        Model model = new Model();

        // Configure the reporting engine to embed inline error messages
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report and capture the success flag
        bool success = engine.BuildReport(doc, model, "model");

        // Output the success flag – expected to be false because of the error above
        Console.WriteLine($"Report build success: {success}");

        // Save the generated report (contains the inline error message)
        doc.Save(outputPath);
    }
}
