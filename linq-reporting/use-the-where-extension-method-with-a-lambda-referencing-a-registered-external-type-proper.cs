using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    public static class FilterSettings
    {
        // External static property used for filtering
        public static string Category { get; set; } = "Electronics";
    }

    public class Product
    {
        public string Name { get; set; } = "";
        public string Category { get; set; } = "";
        public decimal Price { get; set; }
    }

    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var allProducts = new List<Product>
            {
                new() { Name = "Smartphone", Category = "Electronics", Price = 699.99m },
                new() { Name = "Laptop", Category = "Electronics", Price = 1199.49m },
                new() { Name = "Desk Chair", Category = "Furniture", Price = 149.00m },
                new() { Name = "Coffee Table", Category = "Furniture", Price = 89.99m },
                new() { Name = "Headphones", Category = "Electronics", Price = 199.95m }
            };

            // Advanced filtering using Where and external static property
            var filteredProducts = allProducts
                .Where(p => p.Category == FilterSettings.Category)
                .ToList();

            var model = new ReportModel { Products = filteredProducts };

            // Create template document programmatically
            string templatePath = "ReportTemplate.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Product Report - Category: " + FilterSettings.Category);
            builder.Writeln();
            builder.Writeln("<<foreach [p in Products]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("Category: <<[p.Category]>>");
            builder.Writeln("Price: $<<[p.Price]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load the template for reporting
            Document doc = new Document(templatePath);

            // Configure reporting engine
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;

            // Build the report
            bool success = engine.BuildReport(doc, model, "model");

            // Save the generated report
            string outputPath = "ProductReport.docx";
            doc.Save(outputPath);

            // Simple verification output (no interactive input)
            Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
        }
    }
}
