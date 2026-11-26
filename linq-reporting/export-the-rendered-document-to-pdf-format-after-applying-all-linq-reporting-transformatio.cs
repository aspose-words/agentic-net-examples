using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace LinqReportingPdfExport
{
    // Sample data model classes
    public class Order
    {
        public string CustomerName { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any required encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output files.
            const string templatePath = "template.docx";
            const string pdfOutputPath = "report.pdf";

            // -----------------------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title.
            builder.Writeln("Order Report");
            builder.Writeln();

            // Customer name placeholder.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Begin foreach block.
            builder.Writeln("<<foreach [item in order.Items]>>");

            // Table header.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row (will be repeated for each item).
            builder.InsertCell();
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.EndRow();

            // End table and foreach block.
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report using LINQ Reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data to bind.
            Order order = new Order
            {
                CustomerName = "John Doe",
                Items = new List<Item>
                {
                    new Item { Index = 1, Name = "Item A" },
                    new Item { Index = 2, Name = "Item B" },
                    new Item { Index = 3, Name = "Item C" }
                }
            };

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, order, "order");

            // -----------------------------------------------------------------
            // 3. Export the rendered document to PDF.
            // -----------------------------------------------------------------
            reportDoc.Save(pdfOutputPath, SaveFormat.Pdf);
        }
    }
}
