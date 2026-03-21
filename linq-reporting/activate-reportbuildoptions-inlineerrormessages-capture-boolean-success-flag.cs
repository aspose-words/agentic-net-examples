using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineExample
{
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load a template document that contains reporting tags.
            Document template = new Document("Template.docx");

            // Prepare a simple data source.
            var dataSource = new Person
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
            bool success = engine.BuildReport(template, dataSource);

            // Output the result of the build operation.
            Console.WriteLine($"Report build successful: {success}");

            // Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
