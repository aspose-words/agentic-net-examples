using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create sample JSON data and deserialize it to a strongly‑typed model.
        // -----------------------------------------------------------------
        string jsonPath = Path.Combine(workDir, "sample.json");
        File.WriteAllText(jsonPath,
            @"{
                ""Categories"": [
                    { ""Name"": ""Fruits"", ""Items"": [""Apple"", ""Banana"", ""Orange""] },
                    { ""Name"": ""Vegetables"", ""Items"": [""Carrot"", ""Broccoli""] },
                    { ""Name"": ""Beverages"", ""Items"": [""Tea"", ""Coffee"", ""Juice""] }
                ]
            }");

        // Root object that matches the JSON structure.
        RootModel data = JsonConvert.DeserializeObject<RootModel>(File.ReadAllText(jsonPath))!;

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Create a bulleted list style that will be used for all levels.
        List bulletList = template.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Outer foreach – iterates over categories.
        builder.Writeln("<<foreach [category in Categories]>>");
        // Level 0 – category name.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[category.Name]>>");

        // Inner foreach – iterates over items of the current category.
        builder.Writeln("<<foreach [item in category.Items]>>");
        // Level 1 – item name.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<[item]>>");
        builder.Writeln("<</foreach>>"); // end inner foreach

        builder.Writeln("<</foreach>>"); // end outer foreach

        // Clean up list formatting.
        builder.ListFormat.RemoveNumbers();

        // -----------------------------------------------------------------
        // 3. Generate the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this simple scenario.
        engine.Options = ReportBuildOptions.None;

        // BuildReport overload without a data source name allows direct access to root members.
        bool success = engine.BuildReport(template, data);
        if (!success)
        {
            Console.WriteLine("Report generation failed due to template errors.");
        }

        // -----------------------------------------------------------------
        // 4. Save the resulting document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "NestedBulletedList.docx");
        template.Save(outputPath);
        Console.WriteLine($"Report saved to: {outputPath}");
    }
}

// ---------------------------------------------------------------------
// Data model that mirrors the JSON structure.
// ---------------------------------------------------------------------
public class RootModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
