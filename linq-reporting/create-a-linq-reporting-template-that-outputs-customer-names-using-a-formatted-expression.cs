using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model representing a customer.
    public class Customer
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = string.Empty;
    }

    // Wrapper class that will be passed as the root data source.
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a LINQ Reporting template.
            // Header.
            builder.Writeln("Customer List:");
            // Start a foreach loop over the Customers collection.
            builder.Writeln("<<foreach [c in Customers]>>");
            // Output each customer's name using a formatted expression tag.
            builder.Writeln(" - <<[c.Name]>>");
            // End the foreach block.
            builder.Writeln("<</foreach>>");

            // Prepare sample data.
            ReportModel model = new ReportModel();
            model.Customers.Add(new Customer { Name = "Alice Johnson" });
            model.Customers.Add(new Customer { Name = "Bob Smith" });
            model.Customers.Add(new Customer { Name = "Charlie Davis" });

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "model".
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("LinqReportingOutput.docx");
        }
    }
}
