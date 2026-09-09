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
        // Register code page provider for legacy encodings (required by Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a large data set.
        var model = new ReportModel
        {
            Items = new List<Item>()
        };
        for (int i = 1; i <= 1000; i++)
        {
            model.Items.Add(new Item { Index = i, Name = $"Item {i}" });
        }

        // Create a template document with LINQ Reporting tags.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Enable reflection optimization (caching of generated dynamic types).
        ReportingEngine.UseReflectionOptimization = true;

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Write a heading.
        builder.Writeln("Large Data Set Report");
        builder.Writeln();

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Index: <<[item.Index]>> - Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model used in the foreach loop.
public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
