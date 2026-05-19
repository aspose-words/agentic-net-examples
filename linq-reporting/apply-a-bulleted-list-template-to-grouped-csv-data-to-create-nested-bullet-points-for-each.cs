using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (if needed).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample CSV data.
        // -----------------------------------------------------------------
        string csvPath = "data.csv";
        string[] csvLines =
        {
            "Category,Item",
            "Fruits,Apple",
            "Fruits,Banana",
            "Fruits,Orange",
            "Vegetables,Carrot",
            "Vegetables,Tomato",
            "Vegetables,Potato"
        };
        File.WriteAllLines(csvPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Load CSV into a DataTable.
        // -----------------------------------------------------------------
        DataTable table = new DataTable();
        table.Columns.Add("Category", typeof(string));
        table.Columns.Add("Item", typeof(string));

        foreach (string line in File.ReadAllLines(csvPath).Skip(1))
        {
            string[] parts = line.Split(',');
            if (parts.Length == 2)
                table.Rows.Add(parts[0].Trim(), parts[1].Trim());
        }

        // -----------------------------------------------------------------
        // 3. Group data by category and build a model for the report.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Categories = table.AsEnumerable()
                .GroupBy(r => r.Field<string>("Category"))
                .Select(g => new CategoryGroup
                {
                    Category = g.Key,
                    Items = g.Select(r => r.Field<string>("Item")).ToList()
                })
                .ToList()
        };

        // -----------------------------------------------------------------
        // 4. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Create a bulleted list.
        List bulletList = templateDoc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = bulletList;

        // Outer foreach – categories.
        builder.Writeln("<<foreach [cat in Categories]>>");

        // Category bullet (level 0).
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[cat.Category]>>");

        // Inner foreach – items.
        builder.Writeln("<<foreach [itm in cat.Items]>>");

        // Item bullet (level 1).
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<[itm]>>");

        // Close inner foreach.
        builder.Writeln("<</foreach>>");

        // Reset to level 0 for the next category.
        builder.ListFormat.ListLevelNumber = 0;

        // Close outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 5. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the final document.
        // -----------------------------------------------------------------
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<CategoryGroup> Categories { get; set; } = new();
}

public class CategoryGroup
{
    public string Category { get; set; } = "";
    public List<string> Items { get; set; } = new();
}
