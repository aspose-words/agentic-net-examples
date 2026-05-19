using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingExample
{
    // Data model for a single line item.
    public class OrderItem
    {
        public string Description { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
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
            // 1. Prepare sample data.
            var model = new ReportModel
            {
                Items = new List<OrderItem>
                {
                    new OrderItem { Description = "Apple",  Quantity = 3, UnitPrice = 0.50m },
                    new OrderItem { Description = "Banana", Quantity = 5, UnitPrice = 0.30m },
                    new OrderItem { Description = "Orange", Quantity = 2, UnitPrice = 0.80m }
                }
            };

            // 2. Create a template document programmatically.
            const string templatePath = "Template.docx";
            var builder = new DocumentBuilder();

            builder.Writeln("Invoice");
            builder.Writeln(); // empty line

            // Start the foreach band before the table.
            builder.Writeln("<<foreach [item in Items]>>");

            // Build the table that will be repeated for each item.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell(); builder.Writeln("Item");
            builder.InsertCell(); builder.Writeln("Qty");
            builder.InsertCell(); builder.Writeln("Price");
            builder.InsertCell(); builder.Writeln("Line Total");
            builder.EndRow();

            // Data row – each iteration of the foreach will populate a new row.
            builder.InsertCell(); builder.Writeln("<<[item.Description]>>");
            builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
            builder.InsertCell(); builder.Writeln("<<[item.UnitPrice]>>");
            // Calculate line total directly in the expression tag.
            builder.InsertCell(); builder.Writeln("<<[item.Quantity * item.UnitPrice]>>");
            builder.EndRow();

            // Finish the table and close the foreach band.
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template.
            builder.Document.Save(templatePath);

            // 3. Load the template for report generation.
            var doc = new Document(templatePath);

            // 4. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 5. Save the final report.
            doc.Save("Report.docx");
        }
    }
}
