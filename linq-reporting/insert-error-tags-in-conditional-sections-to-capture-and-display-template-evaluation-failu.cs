using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 28;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings Aspose.Words might need.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Simple placeholder.
        builder.Writeln("Customer: <<[model.Name]>>");

        // Conditional that will succeed.
        builder.Writeln("<<if [model.Age > 20]>>Age is greater than 20.<</if>>");

        // Conditional that references a missing property to trigger an error.
        // The InlineErrorMessages option will insert the error message into the output.
        builder.Writeln("<<if [model.NonExistent]>>This will cause an error.<</if>>");

        // Save the template (optional, shown for clarity).
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Prepare the data model.
        var model = new ReportModel();

        // Configure the reporting engine to inline error messages.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report. The root object name must match the tag prefix used in the template.
        bool success = engine.BuildReport(template, model, "model");

        // Save the generated report.
        const string outputPath = "output.docx";
        template.Save(outputPath);

        // Output the success flag.
        Console.WriteLine($"Report generation success: {success}");
    }
}
