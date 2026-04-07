using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingGroupByExample
{
    // Simple data model that matches the JSON structure.
    public class Product
    {
        public string Name { get; set; } = "";
        public string Category { get; set; } = "";
        public double Value { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Register code page provider for any non‑UTF8 encodings that Aspose.Words might need.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create sample JSON data file.
            // -----------------------------------------------------------------
            string jsonPath = "data.json";
            var sampleProducts = new List<Product>
            {
                new Product { Name = "Apple",  Category = "Fruits",  Value = 1.2 },
                new Product { Name = "Banana", Category = "Fruits",  Value = 0.8 },
                new Product { Name = "Carrot", Category = "Vegetables", Value = 0.5 },
                new Product { Name = "Broccoli", Category = "Vegetables", Value = 0.9 },
                new Product { Name = "Chicken", Category = "Meat", Value = 5.0 }
            };
            // Serialize to JSON (using simple string building to avoid extra dependencies).
            string jsonContent = "[" + string.Join(",", sampleProducts.ConvertAll(p =>
                $"{{\"Name\":\"{p.Name}\",\"Category\":\"{p.Category}\",\"Value\":{p.Value}}}")) + "]";
            File.WriteAllText(jsonPath, jsonContent);

            // -----------------------------------------------------------------
            // 2. Build the template document programmatically.
            // -----------------------------------------------------------------
            string templatePath = "template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title
            builder.Writeln("Products grouped by Category");
            builder.Writeln();

            // LINQ Reporting tags:
            // Outer foreach iterates over groups created by GroupBy on the JSON collection.
            builder.Writeln("<<foreach [g in products.GroupBy(p => p.Category)]>>");
            builder.Writeln("Category: <<[g.Key]>>");
            builder.Writeln("<<foreach [p in g]>>");
            builder.Writeln("- <<[p.Name]>> : <<[p.Value]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and run the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            // The JSON data source is treated as a collection named "products".
            JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

            ReportingEngine engine = new ReportingEngine();
            // Build the report using the root name "products" to match the tags in the template.
            engine.BuildReport(reportDoc, jsonDataSource, "products");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            string outputPath = "output.docx";
            reportDoc.Save(outputPath);
        }
    }
}
