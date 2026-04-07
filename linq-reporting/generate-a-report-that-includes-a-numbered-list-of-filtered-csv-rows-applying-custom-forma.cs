using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Id,Name,Category",
            "1,Apple,A",
            "2,Banana,B",
            "3,Cherry,A",
            "4,Date,B",
            "5,Elderberry,A"
        });

        // 2. Load CSV and filter rows where Category == "A".
        List<Item> allItems = File.ReadAllLines(csvPath)
            .Skip(1) // skip header
            .Select(line => line.Split(','))
            .Select(parts => new Item
            {
                Id = int.Parse(parts[0]),
                Name = parts[1],
                Category = parts[2]
            })
            .Where(item => item.Category == "A")
            .ToList();

        // 3. Prepare the model for the reporting engine.
        ReportModel model = new ReportModel { Items = allItems };

        // 4. Build the template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Create a numbered list for the report items.
        List list = templateDoc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;

        // Restart numbering before the foreach block.
        builder.Writeln("<<restartNum>><<foreach [item in Items]>>");

        // Write each item with custom background for even Ids.
        builder.Writeln(
            "<<if [item.Id % 2 == 0]>>" +
            "<<backColor [\"LightGray\"]>><<[item.Id]>> - <<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.Id % 2 != 0]>>" +
            "<<[item.Id]>> - <<[item.Name]>>" +
            "<</if>>");

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // End the list formatting.
        builder.ListFormat.RemoveNumbers();

        // Save the template.
        templateDoc.Save(templatePath);

        // 5. Load the template and build the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(reportDoc, model, "model");

        // 6. Save the final report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive prompts).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
