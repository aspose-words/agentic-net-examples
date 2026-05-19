using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data.
        var products = new List<Product>
        {
            new Product { Id = 1, Name = "Apple", Quantity = 10 },
            new Product { Id = 2, Name = "Banana", Quantity = 20 },
            new Product { Id = 3, Name = "Cherry", Quantity = 15 }
        };

        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(products, Formatting.Indented));

        // Create the template document with an in‑table list.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Product List:");
        builder.Writeln("<<foreach [p in items]>>");

        // Table header.
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Table row bound to the JSON data.
        builder.InsertCell();
        builder.Writeln("<<[p.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Quantity]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template and bind the JSON data source.
        var document = new Document(templatePath);
        var jsonData = new JsonDataSource(jsonPath);

        var engine = new ReportingEngine();
        engine.BuildReport(document, jsonData, "items");

        // Save the generated report.
        const string reportPath = "Report.docx";
        document.Save(reportPath);
    }
}

// Simple data model matching the JSON structure.
public class Product
{
    public int Id { get; set; } = 0;
    public string Name { get; set; } = "";
    public int Quantity { get; set; } = 0;
}
