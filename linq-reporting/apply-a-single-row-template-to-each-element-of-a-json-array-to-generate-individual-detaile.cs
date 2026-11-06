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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        ReportModel model = new()
        {
            Items =
            {
                new Item { Title = "Item 1", Description = "First item description.", Price = 10.5 },
                new Item { Title = "Item 2", Description = "Second item description.", Price = 20.0 },
                new Item { Title = "Item 3", Description = "Third item description.", Price = 30.75 }
            }
        };

        // Serialize model to JSON file.
        const string jsonPath = "data.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(model, Formatting.Indented));

        // Create the template document programmatically.
        const string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Report of Items");
        builder.Writeln("==============");
        builder.Writeln();

        // Begin foreach over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Title: <<[item.Title]>>");
        builder.Writeln("Description: <<[item.Description]>>");
        builder.Writeln("Price: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new(templatePath);

        // Load JSON data back into the model.
        ReportModel data = JsonConvert.DeserializeObject<ReportModel>(File.ReadAllText(jsonPath))!;

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(doc, data, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Root wrapper class for the JSON data.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item class.
public class Item
{
    public string Title { get; set; } = "";
    public string Description { get; set; } = "";
    public double Price { get; set; }
}
