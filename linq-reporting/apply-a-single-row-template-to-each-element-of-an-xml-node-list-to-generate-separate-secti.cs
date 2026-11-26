using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create sample XML data.
        string xmlPath = Path.Combine(outputDir, "data.xml");
        string xmlContent = @"<root>
    <item>
        <title>First Item</title>
        <description>This is the first item.</description>
    </item>
    <item>
        <title>Second Item</title>
        <description>This is the second item.</description>
    </item>
    <item>
        <title>Third Item</title>
        <description>This is the third item.</description>
    </item>
</root>";
        File.WriteAllText(xmlPath, xmlContent);

        // Load XML into a strongly‑typed model.
        ReportModel model = LoadModelFromXml(xmlPath);

        // Build the template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // foreach over the collection of items.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Title: <<[item.Title]>>");
        builder.Writeln("Description: <<[item.Description]>>");
        // Insert a page break after each item to create separate sections.
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("<</foreach>>");

        // Save the template (optional, for inspection).
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        template.Save(reportPath);
    }

    private static ReportModel LoadModelFromXml(string xmlFilePath)
    {
        XDocument doc = XDocument.Load(xmlFilePath);
        var model = new ReportModel();

        foreach (var itemElem in doc.Root?.Elements("item") ?? new List<XElement>())
        {
            var item = new Item
            {
                Title = (string?)itemElem.Element("title") ?? string.Empty,
                Description = (string?)itemElem.Element("description") ?? string.Empty
            };
            model.Items.Add(item);
        }

        return model;
    }
}

// Public data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}
