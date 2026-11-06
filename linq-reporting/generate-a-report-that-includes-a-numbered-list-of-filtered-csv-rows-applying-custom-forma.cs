using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CsvItem
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
}

public class ReportModel
{
    public List<CsvItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample CSV data.
        const string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Id,Name,Category\n" +
            "1,Apple,A\n" +
            "2,Banana,B\n" +
            "3,Cherry,A\n" +
            "4,Date,B\n" +
            "5,Elderberry,A");

        // Load and filter CSV rows where Category == "A".
        var model = new ReportModel();
        var lines = File.ReadAllLines(csvPath);
        var dataLines = lines.Skip(1); // Skip header.
        int index = 1;
        foreach (var line in dataLines)
        {
            var parts = line.Split(',');
            if (parts.Length < 3) continue;
            var category = parts[2].Trim();
            if (category != "A") continue;

            model.Items.Add(new CsvItem
            {
                Index = index++,
                Name = parts[1].Trim(),
                Category = category
            });
        }

        // Create the LINQ Reporting template.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Filtered Items Report");
        builder.Writeln(""); // Empty line.

        // Begin foreach over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Apply background color to even-indexed rows.
        builder.Writeln(
            "<<if [item.Index % 2 == 0]>>" +
            "<<backColor [\"LightGray\"]>><<[item.Index]>>. <<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.Index % 2 != 0]>> <<[item.Index]>>. <<[item.Name]>> <</if>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        const string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
