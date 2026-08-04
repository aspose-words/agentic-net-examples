using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model matching the template.
    public class Item
    {
        public string Category { get; set; } = "";
        public int Amount { get; set; }
    }

    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public static void Main()
    {
        // Paths for temporary files.
        string templatePath = "template.docx";
        string outputPath = "Report.docx";

        // Build the LINQ Reporting template.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Category Summary Report");
        // Loop over groups of items grouped by Category.
        builder.Writeln("<<foreach [g in Items.GroupBy(i => i.Category)]>>");
        builder.Writeln("Category: <<[g.Key]>>");
        builder.Writeln("Total Amount: <<[g.Sum(i => i.Amount)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template and reload it (required before building the report).
        templateDoc.Save(templatePath);
        var reportDoc = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel();
        model.Items.AddRange(new[]
        {
            new Item { Category = "Food",   Amount = 10 },
            new Item { Category = "Food",   Amount = 20 },
            new Item { Category = "Drink",  Amount = 5 },
            new Item { Category = "Drink",  Amount = 15 },
            new Item { Category = "Other",  Amount = 7 }
        });

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        // The root object name is "model" and must match the name used in BuildReport.
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save(outputPath);
    }
}
