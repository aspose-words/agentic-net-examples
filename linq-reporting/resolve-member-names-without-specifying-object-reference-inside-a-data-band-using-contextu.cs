using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Sample data.
            Order order = new()
            {
                Customer = "John Doe",
                Items = new()
                {
                    new Item { Name = "Apple", Price = 1.20m },
                    new Item { Name = "Banana", Price = 0.80m },
                    new Item { Name = "Cherry", Price = 2.50m }
                }
            };

            // Create the template document.
            const string templatePath = "Template.docx";
            Document template = new();
            DocumentBuilder builder = new(template);

            // Header – explicit root reference.
            builder.Writeln("Order for <<[order.Customer]>>");
            builder.Writeln();

            // Data band – members accessed without object reference.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item: <<[Name]>>   Price: <<[Price]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            template.Save(templatePath);

            // Load the template before building the report.
            Document report = new(templatePath);

            // Build the report using the root object name "order".
            ReportingEngine engine = new();
            engine.BuildReport(report, order, "order");

            // Save the generated report.
            report.Save("Report.docx");
        }
    }

    // Root data model.
    public class Order
    {
        public string Customer { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    // Item model used inside the data band.
    public class Item
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
    }
}
