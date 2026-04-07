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
        // Ensure the code page provider is available (required for some CSV scenarios).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create a sample CSV file with two columns: Category,Item
        const string csvPath = "data.csv";
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

        // 2. Load CSV data and group it by Category.
        //    For simplicity we parse the CSV ourselves; the focus of the example is the LINQ Reporting part.
        var records = File.ReadAllLines(csvPath)
                          .Skip(1) // skip header
                          .Select(line => line.Split(','))
                          .Select(parts => new Record { Category = parts[0], Item = parts[1] })
                          .ToList();

        var grouped = records
            .GroupBy(r => r.Category)
            .Select(g => new CategoryGroup
            {
                Category = g.Key,
                Items = g.Select(r => r.Item).ToList()
            })
            .ToList();

        var model = new ReportModel { Groups = grouped };

        // 3. Build the template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply a bulleted list style to the whole document.
        List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Outer foreach – iterates over each category group.
        builder.Writeln("<<foreach [group in Groups]>>");
        // Category name – level 0 bullet.
        builder.Writeln("<<[group.Category]>>");

        // Switch to level 1 for the inner items.
        builder.ListFormat.ListLevelNumber = 1;

        // Inner foreach – iterates over items inside the current category.
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("<<[item]>>");
        builder.Writeln("<</foreach>>");

        // Return to level 0 before the next outer iteration.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<</foreach>>");

        // 4. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(doc, model, "model");

        // 5. Save the resulting document.
        doc.Save("Report.docx");
    }

    // Simple record representing a CSV row.
    public class Record
    {
        public string Category { get; set; } = string.Empty;
        public string Item { get; set; } = string.Empty;
    }

    // Wrapper for a category and its items – used as a data source for the report.
    public class CategoryGroup
    {
        public string Category { get; set; } = string.Empty;
        public List<string> Items { get; set; } = new();
    }

    // Root model passed to the ReportingEngine.
    public class ReportModel
    {
        public List<CategoryGroup> Groups { get; set; } = new();
    }
}
