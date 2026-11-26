using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // Generate a large XML data file.
        string xmlPath = Path.Combine(outputDir, "data.xml");
        GenerateLargeXml(xmlPath, itemCount: 5000);

        // Load XML data into the model.
        ReportModel model = LoadModelFromXml(xmlPath);

        // Benchmark without reflection optimization.
        ReportingEngine.UseReflectionOptimization = false;
        Document docWithoutOpt = new Document(templatePath);
        ReportingEngine engineWithoutOpt = new ReportingEngine();
        Stopwatch swWithout = Stopwatch.StartNew();
        engineWithoutOpt.BuildReport(docWithoutOpt, model, "model");
        swWithout.Stop();
        string reportWithoutPath = Path.Combine(outputDir, "report_without_optimization.docx");
        docWithoutOpt.Save(reportWithoutPath);

        // Benchmark with reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;
        Document docWithOpt = new Document(templatePath);
        ReportingEngine engineWithOpt = new ReportingEngine();
        Stopwatch swWith = Stopwatch.StartNew();
        engineWithOpt.BuildReport(docWithOpt, model, "model");
        swWith.Stop();
        string reportWithPath = Path.Combine(outputDir, "report_with_optimization.docx");
        docWithOpt.Save(reportWithPath);

        // Output benchmark results.
        Console.WriteLine($"Report without optimization generated in {swWithout.ElapsedMilliseconds} ms.");
        Console.WriteLine($"Report with optimization generated in {swWith.ElapsedMilliseconds} ms.");
        Console.WriteLine($"Outputs saved to: {outputDir}");
    }

    // Creates a Word template with a table that iterates over Items.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Items Report");
        builder.Writeln("<<foreach [item in Items]>>");

        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("<<[item.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    // Generates an XML file with a specified number of Item elements.
    private static void GenerateLargeXml(string path, int itemCount)
    {
        XDocument doc = new XDocument(new XElement("Items"));
        for (int i = 1; i <= itemCount; i++)
        {
            XElement item = new XElement("Item",
                new XElement("Id", i),
                new XElement("Name", $"Item {i}")
            );
            doc.Root?.Add(item);
        }
        doc.Save(path);
    }

    // Loads the XML file into the ReportModel.
    private static ReportModel LoadModelFromXml(string path)
    {
        XDocument doc = XDocument.Load(path);
        ReportModel model = new();

        foreach (XElement elem in doc.Root?.Elements("Item") ?? new List<XElement>())
        {
            Item item = new()
            {
                Id = (int)elem.Element("Id")!,
                Name = (string)elem.Element("Name")!
            };
            model.Items.Add(item);
        }

        return model;
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Simple item class matching the XML structure.
public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
}
