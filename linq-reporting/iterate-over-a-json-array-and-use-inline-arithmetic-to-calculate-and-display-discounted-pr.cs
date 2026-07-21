using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    class Program
    {
        static void Main()
        {
            // Register code page provider for Aspose.Words (required on .NET Core).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample JSON data (an array of products with price and discount).
            string jsonContent = @"
[
    { ""Name"": ""Widget"",      ""Price"": 100.0, ""Discount"": 0.10 },
    { ""Name"": ""Gadget"",      ""Price"":  59.5, ""Discount"": 0.15 },
    { ""Name"": ""Doohickey"",   ""Price"":  23.0, ""Discount"": 0.05 }
]";
            string jsonPath = "products.json";
            File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

            // Create a blank Word document that will serve as the template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // LINQ Reporting template:
            // Iterate over the JSON array (named 'items') and calculate discounted price inline.
            builder.Writeln("<<foreach [item in items]>>");
            builder.Writeln("Product: <<[item.Name]>>");
            builder.Writeln("Original Price: $<<[item.Price]>>");
            builder.Writeln("Discounted Price: $<<[item.Price * (1 - item.Discount)]>>");
            builder.Writeln("<</foreach>>");

            // Load the JSON data as a data source.
            JsonDataSource dataSource = new JsonDataSource(jsonPath);

            // Build the report using the template and the JSON data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the generated report.
            doc.Save("DiscountedPricesReport.docx");
        }
    }
}
