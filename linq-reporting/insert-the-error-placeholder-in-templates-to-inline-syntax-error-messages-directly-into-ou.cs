using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Correct tag – will be replaced with the model's Name property.
            builder.Writeln("Hello <<[model.Name]>>!");

            // Malformed tag – missing closing ">>". This will generate a syntax error.
            builder.Writeln("This line has a malformed tag: <<[model.Name]");

            // Placeholder that will be replaced with the inline error message.
            builder.Writeln("<<error>>");

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The method returns true only if the template was parsed without errors.
            bool success = engine.BuildReport(template, model, "model");

            // Save the resulting document.
            string outputPath = "InlineErrorReport.docx";
            template.Save(outputPath);

            // Output the result to the console.
            Console.WriteLine($"Report generation success: {success}");
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
