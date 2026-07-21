using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingWhereExample
{
    // Simple data model representing an order.
    public class Order
    {
        public int Id { get; set; } = 0;
        public string CustomerName { get; set; } = "";
        public string Status { get; set; } = "";
    }

    // Wrapper class that will be passed as the root data source to the reporting engine.
    public class Model
    {
        public List<Order> Orders { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare sample data.
            var model = new Model();
            model.Orders.Add(new Order { Id = 1, CustomerName = "Alice", Status = "Pending" });
            model.Orders.Add(new Order { Id = 2, CustomerName = "Bob",   Status = "Shipped" });
            model.Orders.Add(new Order { Id = 3, CustomerName = "Carol", Status = "Pending" });
            model.Orders.Add(new Order { Id = 4, CustomerName = "Dave",  Status = "Cancelled" });

            // 2. Create a template document programmatically.
            const string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Orders with status 'Pending':");
            // Use the built‑in Where extension method inside the foreach tag.
            builder.Writeln("<<foreach [order in Orders.Where(o => o.Status == \"Pending\")]>>");
            builder.Writeln("Id: <<[order.Id]>>, Customer: <<[order.CustomerName]>>");
            builder.Writeln("<</foreach>>");

            // 3. Save the template to disk.
            templateDoc.Save(templatePath);

            // 4. Load the template back (required before building the report).
            var doc = new Document(templatePath);

            // 5. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            // No special options are needed for this simple scenario.
            engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
