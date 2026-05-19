using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace InlineErrorMessagesDemo
{
    // Simple data model with a single property.
    public class Model
    {
        public string Existing { get; set; } = "Present";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // The first tag is valid, the second references a missing member and will trigger an error.
            builder.Writeln("Existing value: <<[model.Existing]>>");
            builder.Writeln("Missing value: <<[model.Missing]>>");

            // Initialize the reporting engine with InlineErrorMessages option.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // Build the report using the model as the data source.
            bool success = engine.BuildReport(template, new Model(), "model");

            // Retrieve the resulting document text.
            string resultText = template.GetText();

            // Output the success flag and indicate whether an error message was embedded.
            Console.WriteLine($"BuildReport success flag: {success}");
            Console.WriteLine($"Document contains error message: {resultText.Contains("Error")}");

            // Save the generated document for inspection.
            template.Save("InlineErrorReport.docx");
        }
    }
}
