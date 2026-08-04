using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingArithmeticExample
{
    // Model classes must be public with public properties.
    public class Order
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any legacy encodings Aspose might need.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin the foreach loop that will repeat the table rows.
            builder.Writeln("<<foreach [item in Items]>>");

            // Start the table that will be repeated for each item.
            Table table = builder.StartTable();

            // Table header.
            builder.InsertCell();
            builder.Writeln("Item");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.InsertCell();
            builder.Writeln("Total");
            builder.EndRow();

            // Data row – use LINQ Reporting tags.
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Price]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Quantity]>>");
            builder.InsertCell();
            // Arithmetic expression: price * quantity.
            builder.Writeln("<<[item.Price * item.Quantity]>>");
            builder.EndRow();

            // End the table and the foreach block.
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            var order = new Order
            {
                Items = new List<Item>
                {
                    new Item { Name = "Apple",  Price = 0.50m, Quantity = 4 },
                    new Item { Name = "Banana", Price = 0.30m, Quantity = 6 },
                    new Item { Name = "Orange", Price = 0.80m, Quantity = 3 }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // BuildReport overload that takes a data source name – the template uses "order".
            engine.BuildReport(loadedTemplate, order, "order");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
