using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDateFormatting
{
    // Extension method to format DateTime according to a locale string.
    public static class DateTimeExtensions
    {
        public static string FormatDate(this DateTime date, string locale)
        {
            try
            {
                var culture = new CultureInfo(locale);
                return date.ToString(culture);
            }
            catch
            {
                // Fallback to invariant culture if the locale is invalid.
                return date.ToString(CultureInfo.InvariantCulture);
            }
        }
    }

    // Sample data model.
    public class Order
    {
        public DateTime OrderDate { get; set; } = DateTime.Now;
        public string CustomerName { get; set; } = "John Doe";
        public List<OrderItem> Items { get; set; } = new();

        // Wrapper method that can be called from LINQ Reporting expressions.
        public string FormatDate(string locale) => OrderDate.FormatDate(locale);
    }

    public class OrderItem
    {
        public int Index { get; set; }
        public string ProductName { get; set; } = string.Empty;
        public decimal Price { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var order = new Order
            {
                OrderDate = new DateTime(2023, 12, 25, 14, 30, 0),
                CustomerName = "Alice Smith",
                Items = new List<OrderItem>
                {
                    new OrderItem { Index = 1, ProductName = "Widget", Price = 19.99m },
                    new OrderItem { Index = 2, ProductName = "Gadget", Price = 29.99m }
                }
            };

            // Create a template document with LINQ Reporting tags.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Order Date (en-US): <<[order.FormatDate(\"en-US\")]>>");
            builder.Writeln("Order Date (fr-FR): <<[order.FormatDate(\"fr-FR\")]>>");
            builder.Writeln("");
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.ProductName]>> - $<<[item.Price]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(templatePath);

            // Load the template and build the report.
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, order, "order");

            // Save the generated report.
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
