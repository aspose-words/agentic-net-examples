using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Item
{
    public string Category { get; set; } = "";
    public string Name { get; set; } = "";
    public double Value { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data.
        var items = new List<Item>
        {
            new Item { Category = "Fruits", Name = "Apple",  Value = 1.2 },
            new Item { Category = "Fruits", Name = "Banana", Value = 0.8 },
            new Item { Category = "Vegetables", Name = "Carrot", Value = 0.5 },
            new Item { Category = "Fruits", Name = "Orange", Value = 1.0 },
            new Item { Category = "Vegetables", Name = "Lettuce", Value = 0.7 }
        };
        string jsonPath = "data.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(items, Formatting.Indented));

        // Create the template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Report grouped by Category");
        builder.Writeln("");

        // Outer foreach iterates over groups created by GroupBy.
        builder.Writeln("<<foreach [g in items.GroupBy(i => i.Category)]>>");
        builder.Writeln("Category: <<[g.Key]>>");
        builder.Writeln("");

        // Inner foreach iterates over items within the current group.
        builder.Writeln("<<foreach [item in g]>>");
        builder.Writeln("- <<[item.Name]>> : <<[item.Value]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Load JSON data as a data source.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report using the data source named "items".
        var engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "items");

        // Save the generated report.
        doc.Save("ReportOutput.docx");
    }
}
