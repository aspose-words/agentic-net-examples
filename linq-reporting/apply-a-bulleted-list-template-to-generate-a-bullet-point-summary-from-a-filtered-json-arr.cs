using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- 1. Prepare sample JSON data ----------
        string jsonPath = "people.json";
        File.WriteAllText(jsonPath,
            @"[
                { ""Title"": ""Report Q1"", ""Category"": ""Important"" },
                { ""Title"": ""Team Meeting"", ""Category"": ""Routine"" },
                { ""Title"": ""Budget Review"", ""Category"": ""Important"" },
                { ""Title"": ""Holiday Schedule"", ""Category"": ""Routine"" }
            ]");

        // Deserialize JSON into a list of Item objects.
        List<Item> allItems = JsonConvert.DeserializeObject<List<Item>>(File.ReadAllText(jsonPath)) ?? new List<Item>();

        // Filter the items – keep only those with Category == "Important".
        List<Item> filteredItems = allItems.Where(i => i.Category == "Important").ToList();

        // Wrap the filtered collection in a model that will be used as the data source.
        ReportModel model = new ReportModel { Items = filteredItems };

        // ---------- 2. Create the LINQ Reporting template ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a heading.
        builder.Writeln("Important Items Summary:");

        // Create a bulleted list using the built‑in list template.
        Aspose.Words.Lists.List bulletList = templateDoc.Lists.Add(Aspose.Words.Lists.ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</foreach>>");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the template to disk (required before BuildReport according to the rules).
        string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the saved template.
        Document loadedTemplate = new Document(templatePath);

        // ---------- 3. Build the report ----------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // ---------- 4. Save the final document ----------
        string outputPath = "BulletSummary.docx";
        loadedTemplate.Save(outputPath);

        // The program finishes here; no user interaction is required.
    }
}

// Public data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item representation.
public class Item
{
    public string Title { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
}
