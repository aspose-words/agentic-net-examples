using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Valid tag – will be replaced with the value of ReportModel.Name.
            builder.Writeln("Customer: <<[model.Name]>>");

            // Invalid tag – the property does not exist on ReportModel.
            // With InlineErrorMessages enabled the engine will insert "<<error>>" at this location.
            builder.Writeln("Missing property: <<[model.NonExistent]>>");

            // Prepare the data source.
            ReportModel model = new ReportModel { Name = "John Doe" };

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report. The boolean indicates whether parsing succeeded.
            bool success = engine.BuildReport(template, model, "model");

            Console.WriteLine($"Report build success: {success}");

            // Save the generated document. It will contain "<<error>>" where the syntax error occurred.
            template.Save("ReportWithInlineErrors.docx");
        }
    }
}
