using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model classes
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Customer
    {
        public string Name { get; set; } = string.Empty;
    }

    class Program
    {
        static void Main()
        {
            // Create sample data
            var model = new ReportModel();
            model.Customers.Add(new Customer { Name = "Alice Johnson" });
            model.Customers.Add(new Customer { Name = "Bob Smith" });
            model.Customers.Add(new Customer { Name = "Charlie Davis" });

            // Create a new blank document and a builder to insert LINQ Reporting tags
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Title
            builder.Writeln("Customer Names:");
            // Begin foreach loop over the Customers collection
            builder.Writeln("<<foreach [c in model.Customers]>>");
            // Output each customer's name on a new line, using a formatted expression tag
            builder.Writeln(" - <<[c.Name]>>");
            // End foreach loop
            builder.Writeln("<</foreach>>");

            // Build the report using the ReportingEngine
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            const string outputPath = "LinqReportingOutput.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Report generated and saved to '{outputPath}'.");
        }
    }
}
