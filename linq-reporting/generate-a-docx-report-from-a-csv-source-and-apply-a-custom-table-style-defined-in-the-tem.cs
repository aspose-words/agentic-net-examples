using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample CSV data
        string csvPath = "data.csv";
        File.WriteAllText(csvPath, "Name,Value\nAlice,100\nBob,200\nCharlie,300");

        // Load CSV into model
        List<Record> records = LoadCsv(csvPath);

        // Create template document with LINQ Reporting tags and custom table style
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Prepare the root data model
        ReportModel model = new() { Records = records };

        // Build the report
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }

    private static List<Record> LoadCsv(string path)
    {
        var list = new List<Record>();
        string[] lines = File.ReadAllLines(path);
        for (int i = 1; i < lines.Length; i++) // Skip header
        {
            string[] parts = lines[i].Split(',');
            if (parts.Length >= 2)
            {
                list.Add(new Record
                {
                    Name = parts[0],
                    Value = parts[1]
                });
            }
        }
        return list;
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Add a title
        builder.Writeln("Report generated from CSV data");
        builder.Writeln();

        // Insert foreach tag
        builder.Writeln("<<foreach [rec in Records]>>");

        // Start a table for each record
        builder.StartTable();

        // First cell - Name
        builder.InsertCell();
        builder.Writeln("<<[rec.Name]>>");

        // Second cell - Value
        builder.InsertCell();
        builder.Writeln("<<[rec.Value]>>");

        // End the row and table
        builder.EndRow();
        builder.EndTable();

        // Close foreach tag
        builder.Writeln("<</foreach>>");

        // Define a custom table style (simple placeholder)
        Style tableStyle = doc.Styles.Add(StyleType.Table, "CustomTableStyle");
        tableStyle.Font.Name = "Arial";
        tableStyle.Font.Size = 10;

        // Apply the custom style to the template table
        Table? templateTable = doc.GetChild(NodeType.Table, 0, true) as Table;
        if (templateTable != null)
        {
            templateTable.Style = tableStyle;
        }

        // Save the template
        doc.Save(path);
    }
}

public class ReportModel
{
    public List<Record> Records { get; set; } = new();
}

public class Record
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}
