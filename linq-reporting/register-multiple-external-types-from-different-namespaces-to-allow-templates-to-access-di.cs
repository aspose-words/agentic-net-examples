using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ExternalTypesDemo
{
    // Namespace A
    namespace NamespaceA
    {
        public class Person
        {
            public string Name { get; set; } = "";
            public int Age { get; set; }
        }
    }

    // Namespace B
    namespace NamespaceB
    {
        public class Company
        {
            public string Name { get; set; } = "";
            public string Location { get; set; } = "";
        }
    }

    // Wrapper root object
    public class ReportData
    {
        public List<NamespaceA.Person> Persons { get; set; } = new();
        public List<NamespaceB.Company> Companies { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var data = new ReportData
            {
                Persons =
                {
                    new NamespaceA.Person { Name = "Alice", Age = 30 },
                    new NamespaceA.Person { Name = "Bob", Age = 45 }
                },
                Companies =
                {
                    new NamespaceB.Company { Name = "Acme Corp", Location = "New York" },
                    new NamespaceB.Company { Name = "Globex Inc", Location = "London" }
                }
            };

            // Create template document
            string templatePath = "Template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("=== Persons ===");
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            builder.Writeln();
            builder.Writeln("=== Companies ===");
            builder.Writeln("<<foreach [c in Companies]>>");
            builder.Writeln("Company: <<[c.Name]>>, Location: <<[c.Location]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load template for reporting
            Document doc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report
            engine.BuildReport(doc, data, "data");

            // Save the generated report
            string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Indicate completion (no interactive input)
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
