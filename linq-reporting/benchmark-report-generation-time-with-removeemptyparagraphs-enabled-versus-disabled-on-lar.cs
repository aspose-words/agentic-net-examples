using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a large data source.
        ReportModel model = new()
        {
            Items = GenerateItems(5000)
        };

        // Create the LINQ Reporting template programmatically.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Benchmark without RemoveEmptyParagraphs.
        BenchmarkReport(templatePath, model, ReportBuildOptions.None, "Report_WithoutRemove.docx");

        // Benchmark with RemoveEmptyParagraphs enabled.
        BenchmarkReport(templatePath, model, ReportBuildOptions.RemoveEmptyParagraphs, "Report_WithRemove.docx");
    }

    // Generates a list of items where half have an empty Description.
    private static List<Item> GenerateItems(int count)
    {
        var list = new List<Item>(count);
        for (int i = 0; i < count; i++)
        {
            list.Add(new Item
            {
                Name = $"Item {i + 1}",
                Description = i % 2 == 0 ? $"Description {i + 1}" : string.Empty
            });
        }
        return list;
    }

    // Creates a simple template that uses LINQ Reporting tags.
    private static void CreateTemplate(string fileName)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Write the Name field.
        builder.Writeln("Name: <<[item.Name]>>");
        // Write the Description field – may be empty.
        builder.Writeln("Description: <<[item.Description]>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        doc.Save(fileName);
    }

    // Runs the report generation, measures elapsed time, and saves the result.
    private static void BenchmarkReport(string templatePath, ReportModel model, ReportBuildOptions options, string outputPath)
    {
        // Load a fresh copy of the template for each run.
        Document doc = new(templatePath);

        ReportingEngine engine = new()
        {
            Options = options
        };

        Stopwatch sw = Stopwatch.StartNew();
        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");
        sw.Stop();

        doc.Save(outputPath);
        Console.WriteLine($"{options}: Elapsed = {sw.ElapsedMilliseconds} ms, Output = {outputPath}");
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item used in the report.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}
