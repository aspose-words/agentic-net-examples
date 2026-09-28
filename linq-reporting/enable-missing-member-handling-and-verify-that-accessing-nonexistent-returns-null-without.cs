using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words in some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Sample data model
        var model = new Model { Name = "John Doe" };

        // Create a template document with a valid and a nonexistent member
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Missing: <<[model.Nonexistent]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting
        var template = new Document(templatePath);

        // Create the reporting engine and enable missing‑member handling
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report
        engine.BuildReport(template, model, "model");

        // Save the generated report
        const string outputPath = "output.docx";
        template.Save(outputPath);

        // Output the resulting text to verify that the missing member produced an empty value
        string result = template.GetText();
        Console.WriteLine("Report generated successfully. Extracted text:");
        Console.WriteLine(result);
    }
}

public class Model
{
    public string Name { get; set; } = string.Empty;
}
