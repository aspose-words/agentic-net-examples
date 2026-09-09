using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class Logger
{
    // Simple logger that writes to console and stores messages.
    private static readonly List<string> _messages = new();

    public static void Log(string message)
    {
        _messages.Add(message);
        Console.WriteLine(message);
    }

    public static IReadOnlyList<string> Messages => _messages;
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Each item logs when its Value property is accessed.
public class Item
{
    public string Name { get; set; } = string.Empty;

    private int _value;
    public int Value
    {
        get
        {
            Logger.Log($"Evaluating Value for item '{Name}': {_value}");
            return _value;
        }
        set => _value = value;
    }
}

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Header.
        builder.Writeln("Report generated with expression logging:");
        builder.Writeln();

        // foreach over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>>, Value: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a temporary file.
        const string templatePath = "ReportTemplate.docx";
        template.Save(templatePath);

        // 2. Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alpha", Value = 10 },
                new Item { Name = "Beta", Value = 20 },
                new Item { Name = "Gamma", Value = 30 }
            }
        };

        // 3. Load the template (simulating a separate load step).
        var doc = new Document(templatePath);

        // 4. Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        // No special options needed for logging; the property getters perform logging.
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        const string outputPath = "ReportResult.docx";
        doc.Save(outputPath);

        // Optional: indicate completion.
        Console.WriteLine($"Report generated and saved to '{outputPath}'.");
    }
}
