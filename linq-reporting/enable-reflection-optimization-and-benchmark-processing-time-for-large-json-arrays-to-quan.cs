using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // Generate large JSON data
        ReportData data = GenerateLargeData(10_000);
        string json = JsonConvert.SerializeObject(data);
        // Deserialize back to object (simulating real scenario)
        ReportData deserializedData = JsonConvert.DeserializeObject<ReportData>(json)!;

        // Benchmark without reflection optimization
        ReportingEngine.UseReflectionOptimization = false;
        Document docWithoutOpt = new Document(templatePath);
        var engineWithoutOpt = new ReportingEngine();
        var swWithout = Stopwatch.StartNew();
        engineWithoutOpt.BuildReport(docWithoutOpt, deserializedData, "data");
        swWithout.Stop();
        string outputWithout = Path.Combine(outputDir, "Report_WithoutOpt.docx");
        docWithoutOpt.Save(outputWithout);

        // Benchmark with reflection optimization
        ReportingEngine.UseReflectionOptimization = true;
        Document docWithOpt = new Document(templatePath);
        var engineWithOpt = new ReportingEngine();
        var swWith = Stopwatch.StartNew();
        engineWithOpt.BuildReport(docWithOpt, deserializedData, "data");
        swWith.Stop();
        string outputWith = Path.Combine(outputDir, "Report_WithOpt.docx");
        docWithOpt.Save(outputWith);

        // Output benchmark results
        Console.WriteLine($"Report without reflection optimization: {swWithout.ElapsedMilliseconds} ms");
        Console.WriteLine($"Report with reflection optimization:    {swWith.ElapsedMilliseconds} ms");
        Console.WriteLine($"Outputs saved to: {outputDir}");
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Simple header with generation time
        builder.Writeln($"Report generated at <<[ReportGenerated]>>");

        // Foreach block for items
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>, Value: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    private static ReportData GenerateLargeData(int count)
    {
        var items = new List<Item>(count);
        for (int i = 1; i <= count; i++)
        {
            items.Add(new Item
            {
                Id = i,
                Name = $"Item {i}",
                Value = i * 0.1
            });
        }

        return new ReportData
        {
            Items = items,
            ReportGenerated = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        };
    }
}

public class ReportData
{
    public List<Item> Items { get; set; } = new();
    public string ReportGenerated { get; set; } = string.Empty;
}

public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public double Value { get; set; }
}
