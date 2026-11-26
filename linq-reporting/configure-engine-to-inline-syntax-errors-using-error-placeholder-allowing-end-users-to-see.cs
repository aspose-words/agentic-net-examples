using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Sample property used in the template.
    public string Name { get; set; } = "John Doe";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create a template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Correct tag – will be replaced with the model's Name.
        builder.Writeln("Hello <<[model.Name]>>!");

        // Intentional error – references a non‑existent property.
        // With InlineErrorMessages enabled the engine will insert <<error>> here.
        builder.Writeln("This will cause an error <<[model.NonExisting]>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Prepare the data model.
        Model model = new Model();

        // Configure the reporting engine to inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The returned flag indicates overall success.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Output the result status.
        Console.WriteLine($"Report build success: {success}");
        Console.WriteLine($"Generated document saved to: {outputPath}");
    }
}
