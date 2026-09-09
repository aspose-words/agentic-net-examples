using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingArithmeticExample
{
    // Data model for a line item.
    public class OrderItem
    {
        public string Name { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<OrderItem> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document with LINQ Reporting tags.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Invoice");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item: <<[item.Name]>>");
            builder.Writeln("Price: $<<[item.Price]>>");
            builder.Writeln("Quantity: <<[item.Quantity]>>");
            // Arithmetic expression: price * quantity.
            builder.Writeln("Total: $<<[item.Price * item.Quantity]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(templatePath);

            // 2. Load the template for report generation.
            var reportDoc = new Document(templatePath);

            // 3. Prepare sample data.
            var model = new ReportModel
            {
                Items = new List<OrderItem>
                {
                    new() { Name = "Apple",  Price = 0.50m, Quantity = 4 },
                    new() { Name = "Banana", Price = 0.30m, Quantity = 6 },
                    new() { Name = "Cherry", Price = 1.20m, Quantity = 2 }
                }
            };

            // 4. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // 5. Save the generated report.
            reportDoc.Save("Report.docx");
        }
    }
}
