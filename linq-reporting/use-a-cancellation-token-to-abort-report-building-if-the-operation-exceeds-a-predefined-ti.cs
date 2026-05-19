using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for Table type

namespace LinqReportingCancellationDemo
{
    // Root data model for the report.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";

        public List<Item> Items { get; set; } = new()
        {
            new Item { Name = "Apple", Quantity = 3 },
            new Item { Name = "Banana", Quantity = 5 },
            new Item { Name = "Cherry", Quantity = 7 }
        };
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a simple template document with LINQ Reporting tags.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            CreateTemplate(templatePath);

            // 2. Load the template.
            Document doc = new Document(templatePath);

            // 3. Prepare the data source.
            Order order = new Order();

            // 4. Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // 5. Set up a cancellation token that aborts after 2 seconds.
            using CancellationTokenSource cts = new CancellationTokenSource();
            cts.CancelAfter(TimeSpan.FromSeconds(2));
            CancellationToken token = cts.Token;

            // 6. Run the report building in a separate task.
            Task<bool> buildTask = Task.Run(() => engine.BuildReport(doc, order, "order"), token);

            try
            {
                // Wait for either the task to finish or the timeout.
                bool completed = Task.WhenAny(buildTask, Task.Delay(TimeSpan.FromSeconds(2), token)).Result == buildTask;

                if (!completed)
                {
                    Console.WriteLine("Report building timed out and was aborted.");
                    return;
                }

                // If the task completed, save the result.
                string resultPath = Path.Combine(outputDir, "Report.docx");
                doc.Save(resultPath);
                Console.WriteLine($"Report generated successfully: {resultPath}");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Report building was cancelled via the cancellation token.");
            }
            catch (AggregateException ae)
            {
                // Unwrap any exceptions thrown inside the task.
                foreach (var ex in ae.InnerExceptions)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                }
            }
        }

        // Creates a Word document containing LINQ Reporting tags.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Header with customer name.
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Begin foreach loop for items.
            builder.Writeln("<<foreach [item in Items]>>");

            // Table with header and data rows.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Product");
            builder.InsertCell();
            builder.Writeln("Quantity");
            builder.EndRow();

            // Data row (repeated for each item).
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Quantity]>>");
            builder.EndRow();

            // End of table.
            builder.EndTable();

            // Close foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(filePath);
        }
    }
}
