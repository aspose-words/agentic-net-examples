using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Text { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        const string outputDir = "output";
        System.IO.Directory.CreateDirectory(outputDir);

        // Paths for the template and generated reports.
        const string templatePath = "template.docx";
        const string reportWithPath = outputDir + "/Report_WithRemoveEmptyParagraphs.docx";
        const string reportWithoutPath = outputDir + "/Report_WithoutRemoveEmptyParagraphs.docx";

        // -----------------------------------------------------------------
        // 1. Create a large template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // LINQ Reporting tags: iterate over Items and write each item's Text.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Text]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before BuildReport).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare a large data source.
        // -----------------------------------------------------------------
        const int totalItems = 5000;
        ReportModel model = new ReportModel();

        for (int i = 0; i < totalItems; i++)
        {
            // Half of the items have empty text to produce empty paragraphs.
            string text = (i % 2 == 0) ? string.Empty : $"Item #{i}";
            model.Items.Add(new Item { Text = text });
        }

        // -----------------------------------------------------------------
        // 3. Benchmark with RemoveEmptyParagraphs enabled.
        // -----------------------------------------------------------------
        Document docWith = new Document(templatePath);
        ReportingEngine engineWith = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        Stopwatch swWith = Stopwatch.StartNew();
        engineWith.BuildReport(docWith, model, "model");
        swWith.Stop();

        docWith.Save(reportWithPath);
        Console.WriteLine($"Report with RemoveEmptyParagraphs: {swWith.ElapsedMilliseconds} ms");

        // -----------------------------------------------------------------
        // 4. Benchmark with RemoveEmptyParagraphs disabled.
        // -----------------------------------------------------------------
        Document docWithout = new Document(templatePath);
        ReportingEngine engineWithout = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        Stopwatch swWithout = Stopwatch.StartNew();
        engineWithout.BuildReport(docWithout, model, "model");
        swWithout.Stop();

        docWithout.Save(reportWithoutPath);
        Console.WriteLine($"Report without RemoveEmptyParagraphs: {swWithout.ElapsedMilliseconds} ms");
    }
}
