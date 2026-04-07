using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Sample data model
    public class Person
    {
        public string Name { get; set; } = string.Empty;
    }

    // Wrapper class required by ReportingEngine (anonymous types are not supported)
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    // Utility class whose static methods will be used in the template
    public static class MyUtilities
    {
        // Returns a greeting for the supplied name
        public static string Greet(string name) => $"Dear {name}";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for .NET Core
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare output folder
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create the template document programmatically
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("<<[MyUtilities.Greet(p.Name)]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            string templatePath = Path.Combine(outputDir, "template.docx");
            template.Save(templatePath);

            // 2. Load the template (demonstrates the load step)
            Document loadedTemplate = new Document(templatePath);

            // 3. Prepare the data source using a public wrapper class
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice" },
                    new Person { Name = "Bob" },
                    new Person { Name = "Charlie" }
                }
            };

            // 4. Configure the ReportingEngine
            ReportingEngine engine = new ReportingEngine();
            // Register the utility class so its static methods can be used in the template
            engine.KnownTypes.Add(typeof(MyUtilities));

            // 5. Build the report
            engine.BuildReport(loadedTemplate, model, "model");

            // 6. Save the generated report
            string reportPath = Path.Combine(outputDir, "report.docx");
            loadedTemplate.Save(reportPath);

            // Indicate completion
            Console.WriteLine("Report generated at: " + reportPath);
        }
    }
}
