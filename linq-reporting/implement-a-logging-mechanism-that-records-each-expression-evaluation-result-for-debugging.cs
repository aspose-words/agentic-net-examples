using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Index = 1, Name = "Alpha" },
                new Item { Index = 2, Name = "Beta" },
                new Item { Index = 3, Name = "Gamma" }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        // Use the LogHelper to record each expression evaluation.
        builder.Writeln("Index: <<[LogHelper.Log(item.Index)]>>");
        builder.Writeln("Name : <<[LogHelper.Log(item.Name)]>>");
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        // Register the helper type so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(LogHelper));

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("ReportOutput.docx");

        // Output the logged evaluation results.
        Console.WriteLine("Expression Evaluation Log:");
        foreach (var entry in LogHelper.Logs)
        {
            Console.WriteLine(entry);
        }
    }
}

// Simple data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item displayed in the report.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}

// Helper class used in the template to log expression values.
public static class LogHelper
{
    // Stores log entries.
    public static List<string> Logs { get; } = new();

    // Logs the provided value and returns its string representation for the template.
    public static string Log(object? value)
    {
        var text = value?.ToString() ?? "null";
        Logs.Add($"Evaluated: {text}");
        return text;
    }
}
