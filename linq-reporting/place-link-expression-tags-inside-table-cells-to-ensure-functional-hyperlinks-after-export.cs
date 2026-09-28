using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Url = "https://example.com", LinkText = "Example Site" },
                new Item { Url = "https://dotnet.microsoft.com", LinkText = ".NET Home" },
                new Item { Url = "https://github.com", LinkText = "GitHub" }
            }
        };

        // Create a template document programmatically
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Title
        builder.Writeln("Link Table Report");
        builder.Writeln();

        // Header table (single row)
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Link");
        builder.EndRow();
        builder.EndTable();

        // Begin foreach block for items – each iteration creates its own row table
        builder.Writeln("<<foreach [item in Model.Items]>>");
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<link [item.Url] [item.LinkText]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Build the report using LINQ Reporting Engine
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "Model");

        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the generated report
        string outputPath = Path.Combine(outputDir, "LinkReport.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Data model classes
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Url { get; set; } = "";
    public string LinkText { get; set; } = "";
}
