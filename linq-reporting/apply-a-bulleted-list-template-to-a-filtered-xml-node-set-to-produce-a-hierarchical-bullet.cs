using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Prepare sample XML data.
        // -----------------------------------------------------------------
        string xmlContent = @"
<catalog>
    <category name='Fruits'>
        <item>Apple</item>
        <item>Banana</item>
        <item>Cherry</item>
    </category>
    <category name='Vegetables'>
        <item>Carrot</item>
        <item>Broccoli</item>
    </category>
    <category name='Beverages'>
        <item>Tea</item>
        <item>Coffee</item>
    </category>
</catalog>";
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "catalog.xml");
        File.WriteAllText(xmlPath, xmlContent.Trim());

        // -----------------------------------------------------------------
        // 2. Load XML into a strongly‑typed model.
        // -----------------------------------------------------------------
        Catalog model = LoadCatalogFromXml(xmlPath);

        // -----------------------------------------------------------------
        // 3. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin outer foreach over categories.
        builder.Writeln("<<foreach [cat in data.Category]>>");

        // Category name as a top‑level bullet.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("<<[cat.Name]>>");

        // Indent for items.
        builder.ListFormat.ListIndent();

        // Begin inner foreach over items.
        builder.Writeln("<<foreach [itm in cat.Items]>>");
        builder.Writeln("<<[itm]>>");
        builder.Writeln("<</foreach>>");

        // Outdent back to top level.
        builder.ListFormat.ListOutdent();

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Reset list formatting for any following content.
        builder.ListFormat.RemoveNumbers();

        // Save the template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "data");

        // Save the final report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(reportPath);
    }

    // -----------------------------------------------------------------
    // Helper method to load the XML into the model.
    // -----------------------------------------------------------------
    private static Catalog LoadCatalogFromXml(string path)
    {
        XDocument doc = XDocument.Load(path);
        Catalog catalog = new();

        foreach (XElement catElem in doc.Root?.Elements("category") ?? [])
        {
            Category cat = new()
            {
                Name = (string?)catElem.Attribute("name") ?? string.Empty,
                Items = new List<string>()
            };

            foreach (XElement itemElem in catElem.Elements("item"))
            {
                cat.Items.Add(itemElem.Value);
            }

            catalog.Category.Add(cat);
        }

        return catalog;
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class Catalog
{
    public List<Category> Category { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
