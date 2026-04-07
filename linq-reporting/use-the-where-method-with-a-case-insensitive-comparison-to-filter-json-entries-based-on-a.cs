using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingExample
{
    // Data entity representing a product.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
    }

    // Wrapper model passed to the reporting engine.
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample JSON data.
            string jsonPath = "data.json";
            File.WriteAllText(jsonPath,
                @"[
                    { ""Name"": ""Phone"",      ""Category"": ""Electronics"" },
                    { ""Name"": ""Shirt"",      ""Category"": ""Clothing"" },
                    { ""Name"": ""Laptop"",     ""Category"": ""electronics"" },
                    { ""Name"": ""Book"",       ""Category"": ""Books"" }
                ]");

            // 2. Load JSON into a list of Product objects.
            List<Product> allProducts = JsonConvert.DeserializeObject<List<Product>>(File.ReadAllText(jsonPath))
                                        ?? new List<Product>();

            // 3. Filter using case‑insensitive comparison.
            List<Product> filteredProducts = allProducts
                .Where(p => p.Category.Equals("electronics", StringComparison.OrdinalIgnoreCase))
                .ToList();

            // 4. Prepare the model for the reporting engine.
            ReportModel model = new()
            {
                Products = filteredProducts
            };

            // 5. Create the template document programmatically.
            Document doc = new();
            DocumentBuilder builder = new(doc);

            builder.Writeln("Products in the \"Electronics\" category:");
            builder.Writeln("<<foreach [p in Products]>>");
            builder.Writeln("- <<[p.Name]>>");
            builder.Writeln("<</foreach>>");

            // 6. Build the report.
            ReportingEngine engine = new();
            engine.BuildReport(doc, model, "model");

            // 7. Save the generated report.
            string outputPath = "Report.docx";
            doc.Save(outputPath);

            // Optional: inform that the process completed.
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
