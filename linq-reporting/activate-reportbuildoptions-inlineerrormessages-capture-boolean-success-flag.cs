using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineExample
{
    class Program
    {
        static void Main()
        {
            // Load a template document that contains reporting tags.
            // Replace the path with the actual location of your template.
            Document template = new Document("Template.docx");

            // Prepare a simple data source. In real scenarios this could be a DataSet,
            // a custom object, JSON, XML, etc.
            var dataSource = new
            {
                Name = "John Doe",
                Age = 30
            };

            // Create a ReportingEngine instance.
            ReportingEngine engine = new ReportingEngine();

            // Activate the InlineErrorMessages option so that syntax errors are
            // inserted into the output document instead of throwing exceptions.
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report and capture the success flag.
            // The returned value is true if the template was parsed without errors.
            bool success = engine.BuildReport(template, dataSource);

            // Output the result of the build operation.
            Console.WriteLine($"Report build successful: {success}");

            // Save the resulting document.
            // Replace the path with the desired output location.
            template.Save("Result.docx");
        }
    }
}
