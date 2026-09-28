using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public string Name { get; set; } = "John Doe";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create template document with a simple LINQ Reporting tag.
        string templatePath = Path.Combine(outputDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Name: <<[model.Name]>>");
        templateDoc.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Prepare the data model.
        var model = new Model();

        // Configure the reporting engine with InlineErrorMessages.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report and capture the success flag.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);

        // Output the success flag.
        Console.WriteLine($"Report build success: {success}");
    }
}
