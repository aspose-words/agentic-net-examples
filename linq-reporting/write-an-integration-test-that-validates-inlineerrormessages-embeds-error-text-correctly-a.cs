using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
}

public class Program
{
    public static void Main(string[] args)
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare template document with a valid tag and an invalid tag to trigger an inline error.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Hello <<[model.Name]>>!");
        builder.Writeln("Missing property: <<[model.Missing]>>"); // This will cause an error.

        // Save and reload the template to satisfy lifecycle rules.
        string templatePath = "template.docx";
        template.Save(templatePath);
        Document loadedTemplate = new Document(templatePath);

        // Create the data model.
        var model = new Person { Name = "John Doe" };

        // Configure the reporting engine to embed inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report.
        bool success = engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        string outputPath = "output.docx";
        loadedTemplate.Save(outputPath);

        // Verify that the success flag is false (because of the missing property) and that the document contains an error message.
        string documentText = loadedTemplate.GetText();
        bool containsErrorMessage = documentText.Contains("Error", StringComparison.OrdinalIgnoreCase);

        Console.WriteLine($"BuildReport success flag: {success}");
        Console.WriteLine($"Document contains inline error message: {containsErrorMessage}");
    }
}
