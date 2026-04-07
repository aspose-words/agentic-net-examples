using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used in the report.
    public class Model
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the code page provider is registered (required for some data sources).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Step 1: Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            // Insert a LINQ Reporting tag that references the model's Name property.
            builder.Writeln("Hello <<[model.Name]>>!");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Step 2: Load the template back (simulating a real-world scenario where the template is stored).
            Document loadedTemplate = new Document(templatePath);

            // Step 3: Define the restricted .NET types before any report generation.
            // Attempting to access members of these types from the template will be blocked.
            ReportingEngine.SetRestrictedTypes(typeof(System.Type), typeof(System.Environment));

            // Step 4: Prepare the data source.
            Model data = new Model();

            // Step 5: Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.BuildReport(loadedTemplate, data, "model");

            // Step 6: Save the generated report.
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);

            // Inform the user that the process completed.
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(reportPath)}");
        }
    }
}
