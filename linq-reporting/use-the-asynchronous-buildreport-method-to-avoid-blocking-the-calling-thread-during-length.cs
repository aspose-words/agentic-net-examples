using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Required for Table type

namespace AsposeWordsLinqReportingAsync
{
    // Simple data model for the report.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public List<Item> Items { get; set; } = new()
        {
            new Item { Name = "Apple", Quantity = 3 },
            new Item { Name = "Banana", Quantity = 5 },
            new Item { Name = "Orange", Quantity = 2 }
        };
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    class Program
    {
        static async Task Main()
        {
            // Register code page provider (required by Aspose.Words on some platforms).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write a title.
            builder.Writeln("Order Report");
            builder.Writeln();

            // Insert a placeholder for the customer's name.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Begin a foreach block to list items.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a simple table header.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Item");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.EndRow();

            // Table row for each item.
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Quantity]>>");
            builder.EndRow();

            builder.EndTable();

            // End the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report asynchronously.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Sample data source.
            Order order = new Order();

            // Run the potentially long BuildReport operation on a background thread.
            bool success = await Task.Run(() => engine.BuildReport(reportDoc, order, "order"));

            // Optionally handle the success flag (relevant when InlineErrorMessages option is used).
            if (!success)
            {
                Console.WriteLine("Report generation completed with errors.");
            }

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
