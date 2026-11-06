using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingCultureExample
{
    // Root model containing a collection of items.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Individual item with a formatted price property.
    public class Item
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }

        // Formats the price using the current culture.
        public string FormattedPrice => Price.ToString("C", CultureInfo.CurrentCulture);
    }

    public class Program
    {
        public static void Main()
        {
            // Sample data.
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Apple", Price = 1.23m },
                    new Item { Name = "Banana", Price = 0.99m },
                    new Item { Name = "Cherry", Price = 2.50m }
                }
            };

            // Create the template document.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Items Report");
            builder.Writeln("<<foreach [item in Items]>>");

            // Table header.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            // Table rows.
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.FormattedPrice]>>");
            builder.EndRow();
            builder.EndTable();

            builder.Writeln("<</foreach>>");

            // Save the template.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template and build the report.
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the final report.
            const string reportPath = "Report.docx";
            doc.Save(reportPath);
        }
    }
}
