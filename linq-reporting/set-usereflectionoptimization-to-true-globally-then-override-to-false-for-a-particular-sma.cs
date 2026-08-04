using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Root model that contains both the large Order data and the small data.
    public class ReportModel
    {
        public Order Order { get; set; } = new Order();
        public SmallData Small { get; set; } = new SmallData();
    }

    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    public class SmallData
    {
        public string Message { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Set reflection optimization globally.
            ReportingEngine.UseReflectionOptimization = true;

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Order section.
            builder.Writeln("Order Report:");
            builder.Writeln("Customer: <<[model.Order.CustomerName]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in model.Order.Items]>>");
            builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
            builder.Writeln("<</foreach>>");

            // Small data section.
            builder.Writeln();
            builder.Writeln("Small Data Section:");
            builder.Writeln("Message: <<[model.Small.Message]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Build a report using the large data source (global optimization enabled).
            // -----------------------------------------------------------------
            // Load the template.
            Document docLarge = new Document(templatePath);

            // Prepare a large data source.
            var largeModel = new ReportModel
            {
                Order = new Order
                {
                    CustomerName = "John Doe",
                    Items = new List<Item>
                    {
                        new Item { Name = "Apple", Price = 1.20 },
                        new Item { Name = "Banana", Price = 0.80 },
                        new Item { Name = "Cherry", Price = 2.50 }
                    }
                },
                Small = new SmallData { Message = "This is the default small message." }
            };

            // Build the report.
            ReportingEngine engineLarge = new ReportingEngine();
            engineLarge.BuildReport(docLarge, largeModel, "model");

            // Save the generated report.
            docLarge.Save("ReportLarge.docx");

            // -----------------------------------------------------------------
            // 3. Override reflection optimization for a small data source.
            // -----------------------------------------------------------------
            ReportingEngine.UseReflectionOptimization = false;

            // Load the template again for the small report.
            Document docSmall = new Document(templatePath);

            // Prepare a small data source (order has no items).
            var smallModel = new ReportModel
            {
                Order = new Order
                {
                    CustomerName = "Small Customer",
                    Items = new List<Item>() // Empty collection.
                },
                Small = new SmallData { Message = "Hello from the small data source!" }
            };

            // Build the report with optimization disabled.
            ReportingEngine engineSmall = new ReportingEngine();
            engineSmall.BuildReport(docSmall, smallModel, "model");

            // Save the generated small report.
            docSmall.Save("ReportSmall.docx");

            // Reset the global setting if further processing is needed.
            ReportingEngine.UseReflectionOptimization = true;
        }
    }
}
