using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model representing a person.
    public class Person
    {
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
    }

    // Wrapper model that holds a collection of persons.
    public class Model
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a foreach loop that iterates over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            // Concatenate first and last name with a space using an expression tag.
            builder.Writeln("<<[p.FirstName + \" \" + p.LastName]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to a local file.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            Model data = new Model
            {
                Persons = new List<Person>
                {
                    new Person { FirstName = "John", LastName = "Doe" },
                    new Person { FirstName = "Jane", LastName = "Smith" },
                    new Person { FirstName = "Alice", LastName = "Johnson" }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple scenario.
            engine.Options = ReportBuildOptions.None;

            // BuildReport using the model as the root object (no root name needed).
            engine.BuildReport(report, data);

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
            report.Save(outputPath);

            // Indicate completion (optional console output, not required for interaction).
            Console.WriteLine("Report generated successfully at: " + outputPath);
        }
    }
}
