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
        // Register code page provider for possible CSV encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create a sample CSV file.
        string csvPath = "sample.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Category,Item",
            "Fruits,Apple",
            "Fruits,Banana",
            "Vegetables,Carrot",
            "Vegetables,Tomato"
        });

        // 2. Load CSV data and transform it into a hierarchical model.
        ReportModel model = BuildReportModelFromCsv(csvPath);

        // 3. Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Create a bulleted list that will be used for both levels.
        List bulletList = template.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Outer foreach – iterate over categories.
        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.ListFormat.ListLevelNumber = 0; // first‑level bullet
        builder.Writeln("<<[cat.Name]>>");

        // Inner foreach – iterate over items belonging to the current category.
        builder.Writeln("<<foreach [it in cat.Items]>>");
        builder.ListFormat.ListLevelNumber = 1; // second‑level bullet
        builder.Writeln("<<[it.Name]>>");
        builder.Writeln("<</foreach>>"); // end inner foreach

        builder.Writeln("<</foreach>>"); // end outer foreach

        // Remove list formatting after the loops so subsequent paragraphs are normal.
        builder.ListFormat.RemoveNumbers();

        // 4. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(template, model, "model");

        // 5. Save the generated document.
        template.Save("Report.docx");
    }

    // Reads the CSV file and creates a hierarchical model suitable for the report.
    private static ReportModel BuildReportModelFromCsv(string csvPath)
    {
        var categories = new Dictionary<string, Category>(StringComparer.OrdinalIgnoreCase);

        foreach (var line in File.ReadLines(csvPath).Skip(1)) // skip header
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;

            var parts = line.Split(',');
            if (parts.Length != 2)
                continue;

            string catName = parts[0].Trim();
            string itemName = parts[1].Trim();

            if (!categories.TryGetValue(catName, out var category))
            {
                category = new Category { Name = catName };
                categories[catName] = category;
            }

            category.Items.Add(new Item { Name = itemName });
        }

        return new ReportModel { Categories = categories.Values.ToList() };
    }
}

// Root object passed to the reporting engine.
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

// Represents a category that contains multiple items.
public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Simple item class.
public class Item
{
    public string Name { get; set; } = string.Empty;
}
