using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Model namespace with a data class.
    namespace Models
    {
        public class Person
        {
            public string Name { get; set; } = "";
            public int Age { get; set; }

            // Static method that can be called from the template.
            public static string GetGreeting()
            {
                return "Hello from Person!";
            }
        }

        // Wrapper class that will be passed as the root data source.
        public class ReportModel
        {
            public List<Person> Persons { get; set; } = new();
        }
    }

    // Utilities namespace with a static helper class.
    namespace Utilities
    {
        public static class MathHelper
        {
            public static int Add(int a, int b) => a + b;
        }
    }

    class Program
    {
        static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert tags that use static members from external types.
            builder.Writeln("Greeting from static method: <<[Person.GetGreeting()]>>");
            builder.Writeln("Result of static addition: <<[MathHelper.Add(5, 7)]>>");
            builder.Writeln("First person name: <<[model.Persons[0].Name]>>");
            builder.Writeln("First person age: <<[model.Persons[0].Age]>>");

            // Save the template to disk (required by the lifecycle rule).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a real scenario).
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data model.
            // -----------------------------------------------------------------
            var model = new Models.ReportModel
            {
                Persons = new List<Models.Person>
                {
                    new Models.Person { Name = "Alice", Age = 30 },
                    new Models.Person { Name = "Bob", Age = 25 }
                }
            };

            // -----------------------------------------------------------------
            // 4. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Register external types from different namespaces.
            engine.KnownTypes.Add(typeof(Models.Person));
            engine.KnownTypes.Add(typeof(Utilities.MathHelper));

            // -----------------------------------------------------------------
            // 5. Build the report.
            // -----------------------------------------------------------------
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            doc.Save(outputPath);

            // The program finishes without waiting for user input.
        }
    }
}
