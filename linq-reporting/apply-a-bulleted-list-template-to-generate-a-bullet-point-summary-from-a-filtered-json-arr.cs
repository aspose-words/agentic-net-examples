using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // ---------- Step 1: Prepare sample JSON data ----------
        string jsonPath = "data.json";
        string jsonContent = @"[
            { ""Title"": ""Review project plan"", ""Category"": ""Important"" },
            { ""Title"": ""Schedule meeting"", ""Category"": ""Other"" },
            { ""Title"": ""Finalize budget"", ""Category"": ""Important"" },
            { ""Title"": ""Update documentation"", ""Category"": ""Other"" }
        ]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // Deserialize and filter the JSON array.
        List<Item> allItems = JsonConvert.DeserializeObject<List<Item>>(File.ReadAllText(jsonPath, Encoding.UTF8)) ?? new List<Item>();
        List<Item> filteredItems = allItems.Where(i => i.Category == "Important").ToList();

        // Wrap the filtered collection in a model that will be used by the reporting engine.
        ReportModel model = new ReportModel { Items = filteredItems };

        // ---------- Step 2: Create the LINQ Reporting template ----------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Create a single‑level bulleted list based on the default bullet template.
        List bulletedList = templateDoc.Lists.AddSingleLevelList(ListTemplate.BulletDefault);

        // Apply the list formatting to the paragraph that will be repeated.
        builder.ListFormat.List = bulletedList;

        // Insert the foreach tag that iterates over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        // Inside the loop write the bullet text.
        builder.Writeln("<<[item.Title]>>");
        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Remove list formatting after the loop to avoid affecting subsequent paragraphs.
        builder.ListFormat.RemoveNumbers();

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Step 3: Load the template and build the report ----------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // ---------- Step 4: Save the generated report ----------
        string outputPath = "BulletSummaryReport.docx";
        reportDoc.Save(outputPath);
    }
}

// Data entity representing a single item in the JSON array.
public class Item
{
    public string Title { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
}

// Wrapper model used as the root data source for the LINQ Reporting engine.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}
