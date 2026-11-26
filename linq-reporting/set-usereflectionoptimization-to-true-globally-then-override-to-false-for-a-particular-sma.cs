using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // Create a simple LINQ Reporting template.
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report of Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Prepare a large data source.
        var largeModel = new ReportModel
        {
            Items = new List<Item>()
        };
        for (int i = 1; i <= 100; i++)
        {
            largeModel.Items.Add(new Item { Name = $"LargeItem{i}" });
        }

        // Build report for the large data source (global optimization is true).
        var largeReportDoc = new Document(templatePath);
        var engineLarge = new ReportingEngine();
        engineLarge.BuildReport(largeReportDoc, largeModel, "model");
        largeReportDoc.Save("LargeReport.docx");

        // Override reflection optimization for a small data source.
        ReportingEngine.UseReflectionOptimization = false;

        // Prepare a small data source.
        var smallModel = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "SmallItemA" },
                new Item { Name = "SmallItemB" }
            }
        };

        // Build report for the small data source (optimization disabled).
        var smallReportDoc = new Document(templatePath);
        var engineSmall = new ReportingEngine();
        engineSmall.BuildReport(smallReportDoc, smallModel, "model");
        smallReportDoc.Save("SmallReport.docx");

        // Reset global setting if needed.
        ReportingEngine.UseReflectionOptimization = true;
    }
}
