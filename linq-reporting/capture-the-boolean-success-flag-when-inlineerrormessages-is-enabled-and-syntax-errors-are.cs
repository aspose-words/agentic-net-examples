using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Sample property used in the template.
    public string Name { get; set; } = "World";
    // No property named MissingProperty – it will trigger an inline error.
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a template document programmatically.
        // The second line references a non‑existent property to cause a
        // runtime error that will be reported inline.
        // -----------------------------------------------------------------
        var builder = new DocumentBuilder();
        builder.Writeln("Hello <<[model.Name]>>!");                         // Correct tag.
        builder.Writeln("This line has an error: <<[model.MissingProperty]>>"); // Missing property.
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Prepare the data model.
        var model = new Model();

        // Configure the reporting engine to embed error messages inline.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. The method returns false because of the error above.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report (it will contain the inline error message).
        doc.Save(outputPath);

        // Output the success flag.
        Console.WriteLine($"Report build success: {success}");
    }
}
