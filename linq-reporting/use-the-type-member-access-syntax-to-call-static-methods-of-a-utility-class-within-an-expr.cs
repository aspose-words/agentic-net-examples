using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Utility class with a static method that will be called from the template.
    public static class MyUtility
    {
        // Returns a greeting message for the supplied name.
        public static string GetGreeting(string name) => $"Hello, {name}!";
    }

    // Simple data model used as the root object for the report.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words operations).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string reportPath = Path.Combine(outputDir, "Report.docx");

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // The expression tag calls the static method MyUtility.GetGreeting,
            // passing the Name property of the root data object.
            builder.Writeln("Greeting: <<[MyUtility.GetGreeting(Name)]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data source.
            Person person = new Person { Name = "World" };

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register the utility type so that its static members can be used in expressions.
            engine.KnownTypes.Add(typeof(MyUtility));

            // Build the report using the LINQ Reporting engine.
            engine.BuildReport(loadedTemplate, person);

            // Save the generated report.
            loadedTemplate.Save(reportPath);
        }
    }
}
