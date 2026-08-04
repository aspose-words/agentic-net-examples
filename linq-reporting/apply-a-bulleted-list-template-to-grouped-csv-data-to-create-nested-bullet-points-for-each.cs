using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create sample CSV data.
        string csvPath = Path.Combine(workDir, "data.csv");
        File.WriteAllLines(csvPath, new[]
        {
            "Category,Item",
            "Fruits,Apple",
            "Fruits,Banana",
            "Fruits,Orange",
            "Vegetables,Carrot",
            "Vegetables,Tomato",
            "Vegetables,Potato"
        });

        // 2. Load CSV and build a hierarchical model.
        ReportModel model = BuildModelFromCsv(csvPath);

        // 3. Create the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(workDir, "template.docx");
        CreateTemplate(templatePath);

        // 4. Load the template and build the report.
        Document templateDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(templateDoc, model, "model");

        // 5. Save the generated report.
        string outputPath = Path.Combine(workDir, "output.docx");
        templateDoc.Save(outputPath);
    }

    // Parses the CSV file and groups items by category.
    private static ReportModel BuildModelFromCsv(string csvFile)
    {
        var groups = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (var line in File.ReadLines(csvFile).Skip(1)) // Skip header
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;

            var parts = line.Split(',');
            if (parts.Length != 2)
                continue;

            string category = parts[0].Trim();
            string item = parts[1].Trim();

            if (!groups.TryGetValue(category, out var list))
            {
                list = new List<string>();
                groups[category] = list;
            }

            list.Add(item);
        }

        var model = new ReportModel
        {
            Groups = groups.Select(g => new CategoryGroup
            {
                Category = g.Key,
                Items = g.Value
            }).ToList()
        };

        return model;
    }

    // Creates a Word document containing the LINQ Reporting tags.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a bulleted list style.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Begin outer foreach over groups.
        builder.Writeln("<<foreach [group in model.Groups]>>");

        // Category line – top‑level bullet.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[group.Category]>>");

        // Begin inner foreach over items.
        builder.Writeln("<<foreach [item in group.Items]>>");

        // Item line – second‑level bullet.
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<[item]>>");

        // End inner foreach.
        builder.Writeln("<</foreach>>");

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Reset list formatting for any following content.
        builder.ListFormat.RemoveNumbers();

        doc.Save(filePath);
    }
}

// Root data model.
public class ReportModel
{
    public List<CategoryGroup> Groups { get; set; } = new();
}

// Represents a category and its items.
public class CategoryGroup
{
    public string Category { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
