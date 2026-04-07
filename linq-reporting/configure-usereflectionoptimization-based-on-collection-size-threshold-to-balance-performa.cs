using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model for the report.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    public class Program
    {
        // Threshold that decides whether to enable reflection optimization.
        private const int CollectionSizeThreshold = 5;

        public static void Main()
        {
            // Prepare sample data.
            Order order = CreateSampleOrder();

            // Decide on reflection optimization based on collection size.
            ReportingEngine.UseReflectionOptimization = order.Items.Count > CollectionSizeThreshold;

            // Create a template document with LINQ Reporting tags.
            string templatePath = "Template.docx";
            CreateTemplateDocument(templatePath);

            // Load the template.
            Document doc = new Document(templatePath);

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, order, "order");

            // Save the generated report.
            string outputPath = "Report.docx";
            doc.Save(outputPath);
        }

        private static Order CreateSampleOrder()
        {
            // Create an order with a variable number of items.
            Order order = new Order
            {
                CustomerName = "Acme Corp."
            };

            // Add sample items.
            order.Items.Add(new Item { Name = "Widget", Quantity = 3 });
            order.Items.Add(new Item { Name = "Gadget", Quantity = 7 });
            order.Items.Add(new Item { Name = "Doohickey", Quantity = 2 });
            // Uncomment the following lines to increase the collection size and trigger optimization.
            //order.Items.Add(new Item { Name = "Thingamajig", Quantity = 5 });
            //order.Items.Add(new Item { Name = "Whatsit", Quantity = 1 });
            //order.Items.Add(new Item { Name = "Contraption", Quantity = 4 });

            return order;
        }

        private static void CreateTemplateDocument(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a title.
            builder.Writeln("Order Report");
            builder.Writeln();

            // Insert customer name tag.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Insert a foreach loop over the Items collection.
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("- <<[item.Name]>> (Quantity: <<[item.Quantity]>>)");
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(filePath);
        }
    }
}
