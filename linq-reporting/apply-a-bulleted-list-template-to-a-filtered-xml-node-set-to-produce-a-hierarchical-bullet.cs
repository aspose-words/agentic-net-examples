using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Root>
    <Category Name=""Fruits"">
        <Item>Apple</Item>
        <Item>Banana</Item>
        <Item>Cherry</Item>
    </Category>
    <Category Name=""Vegetables"">
        <Item>Carrot</Item>
        <Item>Broccoli</Item>
    </Category>
    <Category Name=""Fungi"">
        <Item>Mushroom</Item>
    </Category>
</Root>";
        string xmlPath = "data.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // Load XML and filter categories whose Name starts with 'F'.
        XDocument xDoc = XDocument.Load(xmlPath);
        List<Category> filteredCategories = xDoc.Root?
            .Elements("Category")
            .Where(c => c.Attribute("Name")?.Value?.StartsWith("F") == true)
            .Select(c => new Category
            {
                Name = c.Attribute("Name")?.Value ?? string.Empty,
                Items = c.Elements("Item")
                         .Select(i => i.Value ?? string.Empty)
                         .ToList()
            })
            .ToList() ?? new List<Category>();

        // Build the model for the report.
        ReportModel model = new()
        {
            Categories = filteredCategories
        };

        // Create the template document programmatically.
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("Hierarchical Bullet List Report");
        builder.Writeln();

        // Begin outer foreach over categories.
        builder.Writeln("<<foreach [cat in Categories]>>");
        // Bullet for category name.
        builder.Writeln("• <<[cat.Name]>>");
        // Begin inner foreach over items.
        builder.Writeln("<<foreach [item in cat.Items]>>");
        // Indented bullet for each item.
        builder.Writeln("   • <<[item]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        string templatePath = "Template.docx";
        template.Save(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(template, model, "model");

        // Save the final report.
        string outputPath = "Report.docx";
        template.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
