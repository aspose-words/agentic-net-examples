using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model representing a customer.
    public class Customer
    {
        public string Name { get; set; } = "";
        public string Address { get; set; } = "";
    }

    // Wrapper class that will be used as the root data source for the LINQ Reporting engine.
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1 – Create a template document with LINQ Reporting tags.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Add a title.
            builder.Writeln("Customer Report");
            builder.Writeln();

            // Begin a foreach loop over the Customers collection.
            builder.Writeln("<<foreach [c in Customers]>>");
            builder.Writeln("Name: <<[c.Name]>>");
            builder.Writeln("Address: <<[c.Address]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // Step 2 – Load the template back (required before building the report).
            var loadedDoc = new Document(templatePath);

            // Step 3 – Prepare sample data.
            var model = new ReportModel
            {
                Customers = new List<Customer>
                {
                    new Customer { Name = "Alice Johnson", Address = "123 Maple Street" },
                    new Customer { Name = "Bob Smith", Address = "456 Oak Avenue" },
                    new Customer { Name = "Carol Davis", Address = "789 Pine Road" }
                }
            };

            // Step 4 – Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.BuildReport(loadedDoc, model, "model");

            // Step 5 – Save the generated report.
            var outputPath = "Report.docx";
            loadedDoc.Save(outputPath);

            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
