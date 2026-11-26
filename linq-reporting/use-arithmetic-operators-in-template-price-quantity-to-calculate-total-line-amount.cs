using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingArithmeticExample
{
    // Data model classes
    public class Order
    {
        public List<LineItem> Items { get; set; } = new();
    }

    public class LineItem
    {
        public string Description { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // -------------------------------------------------
            // Create the template document programmatically
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title
            builder.Writeln("Invoice");
            builder.Writeln();

            // Begin foreach loop over Items (must be outside the table)
            builder.Writeln("<<foreach [item in Items]>>");

            // Start the table inside the foreach block
            Table table = builder.StartTable();

            // Header row
            builder.InsertCell();
            builder.Writeln("Description");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.InsertCell();
            builder.Writeln("Total");
            builder.EndRow();

            // Data row (repeated for each item)
            builder.InsertCell();
            builder.Writeln("<<[item.Description]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Price]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Quantity]>>");
            builder.InsertCell();
            // Arithmetic expression: price * quantity
            builder.Writeln("<<[item.Price * item.Quantity]>>");
            builder.EndRow();

            // End the table
            builder.EndTable();

            // End foreach loop
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // Load the template and build the report
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data
            Order order = new Order
            {
                Items = new List<LineItem>
                {
                    new LineItem { Description = "Widget A", Price = 9.99m, Quantity = 3 },
                    new LineItem { Description = "Widget B", Price = 15.50m, Quantity = 2 },
                    new LineItem { Description = "Gadget",   Price = 7.25m, Quantity = 5 }
                }
            };

            // Build the report using LINQ Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, order, "order");

            // Save the generated report
            reportDoc.Save(outputPath);
        }
    }
}
