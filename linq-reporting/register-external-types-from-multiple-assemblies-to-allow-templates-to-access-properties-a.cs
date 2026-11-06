using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExternalTypesExample
{
    // Simulate external assembly A
    namespace LibraryA
    {
        public class Person
        {
            public string Name { get; set; } = "John Doe";
            public int Age { get; set; } = 30;
        }
    }

    // Simulate external assembly B
    namespace LibraryB
    {
        public class Address
        {
            public string City { get; set; } = "New York";
            public string Street { get; set; } = "5th Avenue";
        }
    }

    // Wrapper model used as the root object for the report
    public class ReportModel
    {
        public LibraryA.Person Person { get; set; } = new();
        public LibraryB.Address Address { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure output directory exists
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // Create a template document with LINQ Reporting tags
            string templatePath = Path.Combine(outputDir, "template.docx");
            CreateTemplate(templatePath);

            // Load the template
            Document doc = new(templatePath);

            // Prepare the data model
            ReportModel model = new();

            // Build the report
            ReportingEngine engine = new();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            string reportPath = Path.Combine(outputDir, "report.docx");
            doc.Save(reportPath);

            // Indicate completion (no interactive input)
            Console.WriteLine($"Report generated at: {reportPath}");
        }

        private static void CreateTemplate(string path)
        {
            Document doc = new();
            DocumentBuilder builder = new(doc);

            builder.Writeln("=== LINQ Reporting Example with External Types ===");
            builder.Writeln();

            // Insert tags that reference properties from both external types
            builder.Writeln("Person Name: <<[model.Person.Name]>>");
            builder.Writeln("Person Age: <<[model.Person.Age]>>");
            builder.Writeln("Address City: <<[model.Address.City]>>");
            builder.Writeln("Address Street: <<[model.Address.Street]>>");

            // Save the template
            doc.Save(path);
        }
    }
}
