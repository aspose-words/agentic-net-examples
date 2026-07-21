using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for JSON handling.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample JSON data.
            string jsonPath = "orders.json";
            File.WriteAllText(jsonPath, @"
[
    { ""Id"": 1, ""Quantity"": 2, ""UnitPrice"": 15.5 },
    { ""Id"": 2, ""Quantity"": 5, ""UnitPrice"": 9.99 },
    { ""Id"": 3, ""Quantity"": 1, ""UnitPrice"": 120.0 }
]".Trim());

            // Create a JSON data source.
            JsonDataSource dataSource = new JsonDataSource(jsonPath);

            // Build the template document programmatically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Header.
            builder.Writeln("Order Report");
            builder.Writeln("------------------------------");

            // Iterate over the orders collection.
            builder.Writeln("<<foreach [order in orders]>>");
            builder.Writeln("Order ID: <<[order.Id]>>");
            builder.Writeln("Quantity: <<[order.Quantity]>>");
            builder.Writeln("Unit Price: $<<[order.UnitPrice]>>");
            // Inline arithmetic: calculate line amount.
            builder.Writeln("Line Amount: $<<[order.Quantity * order.UnitPrice]>>");
            builder.Writeln("<</foreach>>");

            builder.Writeln("------------------------------");
            // Inline arithmetic: calculate total amount for all orders.
            builder.Writeln("Total Amount: $<<[orders.Sum(o => o.Quantity * o.UnitPrice)]>>");

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "orders");

            // Save the generated report.
            doc.Save("OrderReport.docx");
        }
    }
}
