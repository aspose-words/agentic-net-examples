using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        List<Item> sourceItems = new()
        {
            new Item { Id = 1, Name = "Apple",  Value = 5 },
            new Item { Id = 2, Name = "Banana", Value = 12 },
            new Item { Id = 3, Name = "Cherry", Value = 8 },
            new Item { Id = 4, Name = "Date",   Value = 15 },
            new Item { Id = 5, Name = "Elderberry", Value = 20 }
        };

        // LINQ: filter Value > 10, project required fields, order by Name.
        List<ItemDto> filtered = sourceItems
            .Where(i => i.Value > 10)
            .Select(i => new ItemDto { Id = i.Id, Name = i.Name, Value = i.Value })
            .OrderBy(i => i.Name)
            .ToList();

        // Wrap the result for the reporting engine.
        ReportModel model = new() { Items = filtered };

        // Create a template document programmatically.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        CreateTemplate(templatePath);

        // Load the template and build the report.
        Document doc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        doc.Save(reportPath);
    }

    // Generates a simple Word template with a foreach tag.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new();
        DocumentBuilder builder = new(doc);

        builder.Writeln("Filtered and Sorted Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>, Value: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// Simple data entity.
public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}

// DTO used in the report.
public class ItemDto
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}

// Wrapper class for the reporting engine.
public class ReportModel
{
    public List<ItemDto> Items { get; set; } = new();
}
