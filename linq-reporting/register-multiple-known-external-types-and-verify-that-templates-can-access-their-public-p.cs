using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple static utility class with a public static property.
    public static class Utils
    {
        public static string Greeting => "Hello from Utils!";
    }

    // Sample data classes.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    public class Company
    {
        public string Name { get; set; } = "Acme Corp";
        public string Location { get; set; } = "New York";
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public Person Person { get; set; } = new();
        public Company Company { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Use tags that reference the model's properties.
            builder.Writeln("Person Name: <<[model.Person.Name]>>");
            builder.Writeln("Person Age: <<[model.Person.Age]>>");
            builder.Writeln("Company Name: <<[model.Company.Name]>>");
            builder.Writeln("Company Location: <<[model.Company.Location]>>");

            // Use a static property from a known external type.
            builder.Writeln("Static Greeting: <<[Utils.Greeting]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the reporting engine.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Register the external types so that the template can access them.
            engine.KnownTypes.Add(typeof(Utils));
            engine.KnownTypes.Add(typeof(Person));
            engine.KnownTypes.Add(typeof(Company));

            // Optional: ensure reflection optimization is enabled (default).
            ReportingEngine.UseReflectionOptimization = true;

            // -----------------------------------------------------------------
            // 3. Build the report using a populated model.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Person = new Person { Name = "Alice Smith", Age = 28 },
                Company = new Company { Name = "Tech Solutions", Location = "Seattle" }
            };

            // Build the report. The root name in the template is "model".
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            loadedTemplate.Save(reportPath);

            // -----------------------------------------------------------------
            // 4. Verify the output by printing the document text to the console.
            // -----------------------------------------------------------------
            Console.WriteLine("=== Generated Report Content ===");
            Console.WriteLine(loadedTemplate.GetText());
        }
    }
}
