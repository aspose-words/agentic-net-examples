using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Paths for template and generated reports.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPathTrue = Path.Combine(outputDir, "report_remove_true.docx");
        string reportPathFalse = Path.Combine(outputDir, "report_remove_false.docx");

        // Create a large data model.
        ReportModel model = new();
        const int itemCount = 5000;
        for (int i = 1; i <= itemCount; i++)
        {
            model.Items.Add(new Item
            {
                Index = i,
                Name = $"Item {i}"
            });
        }

        // Create the LINQ Reporting template programmatically.
        CreateTemplate(templatePath);

        // Load the template once.
        Document templateDoc = new(templatePath);

        // Benchmark with RemoveEmptyParagraphs enabled via ReportingEngine options.
        Document docTrue = (Document)templateDoc.Clone();
        ReportingEngine engineTrue = new();
        engineTrue.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        Stopwatch swTrue = Stopwatch.StartNew();
        bool successTrue = engineTrue.BuildReport(docTrue, model, "model");
        swTrue.Stop();
        docTrue.Save(reportPathTrue);
        Console.WriteLine($"RemoveEmptyParagraphs = true: Success={successTrue}, Time={swTrue.ElapsedMilliseconds} ms");

        // Benchmark with RemoveEmptyParagraphs disabled (default options).
        Document docFalse = (Document)templateDoc.Clone();
        ReportingEngine engineFalse = new();
        // No special options set; default behavior keeps empty paragraphs.
        Stopwatch swFalse = Stopwatch.StartNew();
        bool successFalse = engineFalse.BuildReport(docFalse, model, "model");
        swFalse.Stop();
        docFalse.Save(reportPathFalse);
        Console.WriteLine($"RemoveEmptyParagraphs = false: Success={successFalse}, Time={swFalse.ElapsedMilliseconds} ms");
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        // Insert a heading.
        builder.Writeln("Large Report");

        // Begin foreach tag.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Table row with data placeholders.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        // End table and foreach.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}
