using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

namespace LinqReportingJsonDiscount
{
    // Model representing a product.
    public class Product
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
        public decimal Discount { get; set; } // e.g., 0.15 for 15% discount
    }

    // Wrapper model for the report.
    public class ReportModel
    {
        public List<Product> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by Aspose.Words on some platforms).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // 1. Create sample JSON data.
            string json = @"
            {
                ""Items"": [
                    { ""Name"": ""Laptop"", ""Price"": 1200.00, ""Discount"": 0.10 },
                    { ""Name"": ""Smartphone"", ""Price"": 800.00, ""Discount"": 0.15 },
                    { ""Name"": ""Headphones"", ""Price"": 150.00, ""Discount"": 0.05 }
                ]
            }";

            // 2. Deserialize JSON into the report model.
            ReportModel model = JsonConvert.DeserializeObject<ReportModel>(json) ?? new ReportModel();

            // 3. Create the LINQ Reporting template programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Product Price Report");
            builder.Writeln("Generated on: " + DateTime.Now);
            builder.Writeln();

            // Start foreach loop over Items.
            builder.Writeln("<<foreach [p in Items]>>");

            // Create a table for each product.
            Table table = builder.StartTable();

            // Header row (repeated for simplicity).
            builder.InsertCell();
            builder.Writeln("Product");
            builder.InsertCell();
            builder.Writeln("Original Price");
            builder.InsertCell();
            builder.Writeln("Discounted Price");
            builder.EndRow();

            // Data row.
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[p.Price]>>");
            builder.InsertCell();
            // Inline arithmetic to calculate discounted price.
            builder.Writeln("<<[p.Price * (1 - p.Discount)]>>");
            builder.EndRow();

            builder.EndTable();

            // End foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            string templatePath = "Template.docx";
            template.Save(templatePath);

            // 4. Load the template.
            Document doc = new Document(templatePath);

            // 5. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine("Report generated: " + Path.GetFullPath(outputPath));
        }
    }
}
