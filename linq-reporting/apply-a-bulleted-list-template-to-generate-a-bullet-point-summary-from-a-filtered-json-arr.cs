using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data.
        string jsonPath = "data.json";
        File.WriteAllText(jsonPath,
            @"[
                { ""Text"": ""Apple"" },
                { ""Text"": ""Banana"" },
                { ""Text"": ""Cherry"" },
                { ""Text"": ""Date"" }
            ]");

        // Deserialize JSON into a list of simple objects.
        var rawItems = JsonConvert.DeserializeObject<List<JsonItem>>(File.ReadAllText(jsonPath)) ?? new List<JsonItem>();

        // Filter the items (e.g., keep texts longer than 5 characters).
        var filtered = rawItems
            .Where(i => i.Text != null && i.Text.Length > 5)
            .Select(i => i.Text!)
            .ToList();

        // Wrap the filtered collection in a model class for the reporting engine.
        var model = new ReportModel { Items = filtered };

        // Create a Word document template programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Add a bulleted list to the document.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the model as the data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("BulletSummary.docx");
    }

    // Simple class matching the JSON structure.
    public class JsonItem
    {
        public string? Text { get; set; }
    }

    // Wrapper class used as the root data source for the report.
    public class ReportModel
    {
        public List<string> Items { get; set; } = new();
    }
}
