using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required on some platforms).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // 1. Create sample JSON data.
        string jsonPath = "data.json";
        var sampleData = new List<Item>
        {
            new Item { Name = "Alpha",   Category = "A" },
            new Item { Name = "Beta",    Category = "B" },
            new Item { Name = "Gamma",   Category = "A" },
            new Item { Name = "Delta",   Category = "C" },
            new Item { Name = "Epsilon", Category = "A" }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // 2. Load JSON and filter the array (keep only Category "A").
        var allItems = JsonConvert.DeserializeObject<List<Item>>(File.ReadAllText(jsonPath)) ?? new List<Item>();
        var filteredItems = allItems.Where(i => i.Category == "A").ToList();

        // 3. Prepare the root data model for the report.
        var model = new ReportModel { Items = filteredItems };

        // 4. Build the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Optional title.
        builder.Writeln("Bullet‑point summary (Category = A):");

        // Apply a bulleted list style.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // End list formatting.
        builder.ListFormat.RemoveNumbers();

        // 5. Generate the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 6. Save the resulting document.
        doc.Save("Report.docx");
    }
}
