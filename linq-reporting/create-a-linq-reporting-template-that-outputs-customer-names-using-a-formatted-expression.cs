using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json; // Included as required package
using System.Text; // For encoding provider

namespace LinqReportingExample
{
    // Data model representing a customer.
    public class Customer
    {
        public string Name { get; set; } = "";
    }

    // Wrapper model that holds a collection of customers.
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words operations).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new ReportModel();
            model.Customers.Add(new Customer { Name = "Alice Johnson" });
            model.Customers.Add(new Customer { Name = "Bob Smith" });
            model.Customers.Add(new Customer { Name = "Charlie Brown" });

            // Create a blank document and insert LINQ Reporting tags.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Customer List:");
            // Begin a foreach loop over the Customers collection.
            builder.Writeln("<<foreach [c in model.Customers]>>");
            // Output each customer's name using a formatted expression tag.
            builder.Writeln(" - <<[c.Name]>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Build the report using the model as the root object named "model".
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("CustomerReport.docx");
        }
    }
}
