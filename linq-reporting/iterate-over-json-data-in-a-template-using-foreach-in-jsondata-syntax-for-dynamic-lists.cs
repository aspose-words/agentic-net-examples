using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data
        var products = new List<Product>
        {
            new Product { Name = "Apple", Price = 0.99 },
            new Product { Name = "Banana", Price = 0.59 },
            new Product { Name = "Cherry", Price = 2.49 }
        };
        string jsonPath = "products.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(products, Formatting.Indented));

        // Create a Word template with LINQ Reporting tags
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product List:");
        builder.Writeln("<<foreach [p in jsonData]>>");
        builder.Writeln(" - <<[p.Name]>> : $<<[p.Price]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template
        var reportDoc = new Document(templatePath);

        // Load JSON data source
        var jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "jsonData");

        // Save the generated report
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        // Clean up temporary files (optional)
        // File.Delete(jsonPath);
        // File.Delete(templatePath);
    }

    public class Product
    {
        public string Name { get; set; } = "";
        public double Price { get; set; }
    }
}
