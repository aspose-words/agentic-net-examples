using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Category { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample XML data with a Category attribute.
        string xmlContent = @"
<Products>
    <Product Name='Laptop' Category='Electronics' />
    <Product Name='Desk' Category='Furniture' />
    <Product Name='Smartphone' Category='electronics' />
    <Product Name='Chair' Category='Furniture' />
    <Product Name='Headphones' Category='ELECTRONICS' />
</Products>";

        // Load XML and filter items where Category equals "electronics" (case‑insensitive).
        XDocument xDoc = XDocument.Parse(xmlContent);
        List<Item> filteredItems = xDoc.Root!
            .Elements("Product")
            .Where(e => string.Equals((string?)e.Attribute("Category"), "electronics", StringComparison.OrdinalIgnoreCase))
            .Select(e => new Item
            {
                Name = (string?)e.Attribute("Name") ?? string.Empty,
                Category = (string?)e.Attribute("Category") ?? string.Empty
            })
            .ToList();

        // Prepare the LINQ Reporting template.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Filtered Products (Category = Electronics):");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln(" - <<[item.Name]>> (<<[item.Category]>>)");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Create the root data model.
        ReportModel model = new ReportModel { Items = filteredItems };

        // Build the report using Aspose.Words LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
