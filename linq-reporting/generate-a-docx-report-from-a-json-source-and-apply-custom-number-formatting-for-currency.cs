using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace LinqReportingJsonCurrency
{
    // Wrapper for the report data
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    public class Order
    {
        public string CustomerName { get; set; } = "";
        public decimal TotalAmount { get; set; }

        // Formatted currency string for LINQ Reporting tag
        public string TotalAmountFormatted => TotalAmount.ToString("C", CultureInfo.CurrentCulture);

        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }

        // Formatted currency string for LINQ Reporting tag
        public string PriceFormatted => Price.ToString("C", CultureInfo.CurrentCulture);
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // Create sample JSON data and write it to a file
            // -----------------------------------------------------------------
            string jsonPath = "data.json";
            var sampleData = new ReportModel
            {
                Orders = new List<Order>
                {
                    new()
                    {
                        CustomerName = "Acme Corp",
                        TotalAmount = 1234.56m,
                        Items = new List<Item>
                        {
                            new() { Name = "Widget", Price = 199.99m },
                            new() { Name = "Gadget", Price = 349.99m }
                        }
                    },
                    new()
                    {
                        CustomerName = "Globex Inc",
                        TotalAmount = 7890.12m,
                        Items = new List<Item>
                        {
                            new() { Name = "Doohickey", Price = 499.50m },
                            new() { Name = "Thingamajig", Price = 299.75m }
                        }
                    }
                }
            };
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

            // Load JSON into the model
            var model = JsonConvert.DeserializeObject<ReportModel>(File.ReadAllText(jsonPath))!;

            // -----------------------------------------------------------------
            // Create the template document with LINQ Reporting tags
            // -----------------------------------------------------------------
            string templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Order Report");
            builder.Writeln("==============");
            builder.Writeln();

            // Outer foreach over orders
            builder.Writeln("<<foreach [order in Orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Total: <<[order.TotalAmountFormatted]>>");
            builder.Writeln();

            // Items table for each order
            builder.Writeln("Items:");
            // Begin inner foreach – the whole table will be repeated per item
            builder.Writeln("<<foreach [item in order.Items]>>");

            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Item Name");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.PriceFormatted]>>");
            builder.EndRow();

            builder.EndTable();

            // Close inner foreach
            builder.Writeln("<</foreach>>");

            // Blank line between orders
            builder.Writeln();

            // Close outer foreach
            builder.Writeln("<</foreach>>");

            // Save the template
            doc.Save(templatePath);

            // -----------------------------------------------------------------
            // Build the final report using the template and the model
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report; the root object is "model"
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            string outputPath = "report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
