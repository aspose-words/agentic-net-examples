using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public int Value { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Step 1: Create a simple LINQ Reporting template programmatically.
        var builder = new DocumentBuilder();
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>, Value: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save("Template.docx");

        // Step 2: Load the saved template.
        var doc = new Document("Template.docx");

        // Step 3: Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new() { Name = "Alpha",   Value = 10 },
                new() { Name = "Beta",    Value = 20 },
                new() { Name = "Gamma",   Value = 30 },
                new() { Name = "Delta",   Value = 40 },
                new() { Name = "Epsilon", Value = 50 },
                new() { Name = "Zeta",    Value = 60 },
                new() { Name = "Eta",     Value = 70 }
            }
        };

        // Step 4: Configure reflection optimization based on collection size.
        const int threshold = 5;
        ReportingEngine.UseReflectionOptimization = model.Items.Count > threshold;

        // Step 5: Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Step 6: Save the generated report.
        doc.Save("Report.docx");
    }
}
