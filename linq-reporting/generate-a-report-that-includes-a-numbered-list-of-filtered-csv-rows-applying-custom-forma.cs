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
        // Ensure code page provider is registered (required for some CSV parsing scenarios).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Id,Name,Value",
            "1,Alpha,30",
            "2,Beta,75",
            "3,Gamma,120",
            "4,Delta,45",
            "5,Epsilon,200"
        });

        // 2. Load CSV rows into a list of Item objects.
        List<Item> allItems = new List<Item>();
        foreach (var line in File.ReadAllLines(csvPath).Skip(1)) // Skip header.
        {
            var parts = line.Split(',');
            if (parts.Length != 3) continue;
            if (!int.TryParse(parts[0], out int id)) continue;
            string name = parts[1];
            if (!int.TryParse(parts[2], out int value)) continue;
            allItems.Add(new Item { Id = id, Name = name, Value = value });
        }

        // 3. Filter rows (e.g., keep items with Value > 50).
        List<Item> filteredItems = allItems.Where(i => i.Value > 50).ToList();

        // 4. Prepare the root data model.
        ReportModel model = new ReportModel { Items = filteredItems };

        // 5. Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("Filtered Items Report:");
        builder.Writeln();

        // Create a numbered list.
        List numberedList = template.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;

        // Insert the restartNum tag before the foreach loop in the same numbered paragraph.
        builder.Writeln("<<restartNum>><<foreach [item in Items]>>");

        // Paragraph content with conditional formatting:
        // Items with Value > 100 get a light gray background.
        builder.Writeln(
            "<<if [item.Value > 100]>><<backColor [\"LightGray\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.Value <= 100]>><<[item.Name]>> <</if>>- <<[item.Value]>>");

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the template (optional, for inspection).
        template.Save("template.docx");

        // 6. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(template, model, "model");

        // 7. Save the final report.
        template.Save("Report.docx");
    }
}

// Root data model exposed to the template.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple data entity representing a CSV row.
public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public int Value { get; set; }
}
