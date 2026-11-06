using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Generate large XML data
        string xmlPath = Path.Combine(outputDir, "data.xml");
        ReportModel model = GenerateLargeModel(20000);
        SerializeToXml(model, xmlPath);

        // Load data back from XML
        ReportModel data = DeserializeFromXml(xmlPath);

        // Create template document
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // Benchmark without reflection optimization
        ReportingEngine.UseReflectionOptimization = false;
        var docNoOpt = new Document(templatePath);
        var engineNoOpt = new ReportingEngine();
        var swNoOpt = Stopwatch.StartNew();
        engineNoOpt.BuildReport(docNoOpt, data, "model");
        swNoOpt.Stop();
        string reportNoOptPath = Path.Combine(outputDir, "report_no_opt.docx");
        docNoOpt.Save(reportNoOptPath);

        // Benchmark with reflection optimization
        ReportingEngine.UseReflectionOptimization = true;
        var docOpt = new Document(templatePath);
        var engineOpt = new ReportingEngine();
        var swOpt = Stopwatch.StartNew();
        engineOpt.BuildReport(docOpt, data, "model");
        swOpt.Stop();
        string reportOptPath = Path.Combine(outputDir, "report_opt.docx");
        docOpt.Save(reportOptPath);

        // Output results
        Console.WriteLine($"Report without reflection optimization: {swNoOpt.ElapsedMilliseconds} ms");
        Console.WriteLine($"Report with reflection optimization:    {swOpt.ElapsedMilliseconds} ms");
        Console.WriteLine($"Outputs saved to: {outputDir}");
    }

    private static ReportModel GenerateLargeModel(int itemCount)
    {
        var items = new List<Item>(itemCount);
        for (int i = 1; i <= itemCount; i++)
        {
            items.Add(new Item
            {
                Id = i,
                Name = $"Item {i}"
            });
        }

        return new ReportModel { Items = items };
    }

    private static void SerializeToXml(ReportModel model, string filePath)
    {
        var serializer = new XmlSerializer(typeof(ReportModel));
        using var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        serializer.Serialize(stream, model);
    }

    private static ReportModel DeserializeFromXml(string filePath)
    {
        var serializer = new XmlSerializer(typeof(ReportModel));
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
        return (ReportModel?)serializer.Deserialize(stream) ?? new ReportModel();
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report generated from large XML data:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
}
