using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data
        string jsonPath = "data.json";
        var sampleData = new ReportModel
        {
            Products = new List<Product>
            {
                new Product { Index = 1, Name = "Apple" },
                new Product { Index = 2, Name = "Banana" },
                new Product { Index = 3, Name = "Cherry" }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load JSON into model
        var jsonContent = File.ReadAllText(jsonPath);
        var model = JsonConvert.DeserializeObject<ReportModel>(jsonContent) ?? new ReportModel();

        // Create template document programmatically
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Product List");
        builder.Writeln();

        // Begin foreach loop – the whole table will be repeated for each product
        builder.Writeln("<<foreach [product in Products]>>");

        // Table for a single product
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Writeln("<<[product.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[product.Name]>>");
        builder.EndRow();

        // End table
        builder.EndTable();

        // End foreach loop
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection)
        doc.Save(templatePath);

        // Build the report using the LINQ Reporting engine
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}

// Data model classes
public class ReportModel
{
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
