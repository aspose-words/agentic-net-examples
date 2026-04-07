using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingStaticMethodExample
{
    // Utility class with a static method that will be called from the template.
    public static class GreetingUtility
    {
        public static string GetGreeting(string name)
        {
            return $"Hello, {name}!";
        }
    }

    // Simple data model used as the root object for the report.
    public class PersonModel
    {
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document and a builder to add content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a LINQ Reporting tag that calls the static method.
            // The root object will be referenced as "model", so we can use its members directly.
            // Type member access syntax allows us to call GreetingUtility.GetGreeting(...).
            builder.Writeln("<<[GreetingUtility.GetGreeting(Name)]>>");

            // Prepare the data source.
            PersonModel model = new PersonModel { Name = "Aspose" };

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Register the utility type so its static members can be accessed from the template.
            engine.KnownTypes.Add(typeof(GreetingUtility));

            // Build the report using the model as the root object named "model".
            engine.BuildReport(doc, model, "model");

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // Save the generated document.
            string outputPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
