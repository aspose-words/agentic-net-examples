using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
}

public class LargeDataSource : IEnumerable<Item>
{
    private readonly int _count;

    public LargeDataSource(int count = 10000)
    {
        _count = count;
    }

    public IEnumerator<Item> GetEnumerator()
    {
        for (int i = 1; i <= _count; i++)
        {
            yield return new Item { Id = i, Name = $"Item {i}" };
        }
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}

public class ReportModel
{
    public IEnumerable<Item> Items { get; set; } = new List<Item>();
}

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Report of Large Data Set");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the model with the custom data source.
        ReportModel model = new()
        {
            Items = new LargeDataSource()
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
