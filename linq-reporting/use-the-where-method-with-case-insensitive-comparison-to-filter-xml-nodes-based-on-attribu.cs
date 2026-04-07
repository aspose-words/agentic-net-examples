using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Prepare sample XML data.
        const string xmlContent = @"
<Products>
    <Product Name='Apple' Category='Fruit' />
    <Product Name='Carrot' Category='Vegetable' />
    <Product Name='Banana' Category='fruit' />
    <Product Name='Broccoli' Category='Vegetable' />
</Products>";
        XDocument xDoc = XDocument.Parse(xmlContent);

        // 2. Filter XML nodes using case‑insensitive comparison.
        List<Item> filteredItems = xDoc.Root?
            .Elements("Product")
            .Where(p => string.Equals((string)p.Attribute("Category"), "fruit", StringComparison.OrdinalIgnoreCase))
            .Select(p => new Item
            {
                Name = (string)p.Attribute("Name") ?? "",
                Category = (string)p.Attribute("Category") ?? ""
            })
            .ToList() ?? new List<Item>();

        // 3. Prepare the model for the report.
        ReportModel model = new()
        {
            Items = filteredItems
        };

        // 4. Create a Word template programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Products filtered by Category = \"fruit\" (case‑insensitive):");
        builder.Writeln();
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>, Category: <<[item.Category]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // 5. Load the template and build the report.
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // 6. Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // 7. Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
