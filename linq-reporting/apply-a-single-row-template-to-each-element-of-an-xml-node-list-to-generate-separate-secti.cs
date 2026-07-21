using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample XML data
        string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<root>
    <Item>
        <Title>First Item</Title>
        <Description>This is the first item's description.</Description>
    </Item>
    <Item>
        <Title>Second Item</Title>
        <Description>This is the second item's description.</Description>
    </Item>
    <Item>
        <Title>Third Item</Title>
        <Description>This is the third item's description.</Description>
    </Item>
</root>";

        // Parse XML into a strongly‑typed model
        ReportModel model = new();
        XDocument doc = XDocument.Parse(xmlContent);
        foreach (XElement elem in doc.Root!.Elements("Item"))
        {
            model.Items.Add(new Item
            {
                Title = (string?)elem.Element("Title") ?? string.Empty,
                Description = (string?)elem.Element("Description") ?? string.Empty
            });
        }

        // Create a template document programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Report generated from XML data");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Title: <<[item.Title]>>");
        builder.Writeln("Description: <<[item.Description]>>");
        // Insert a page break after each item to generate separate sections
        builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
        builder.Writeln("<</foreach>>");

        // Save the template (optional, just for demonstration)
        const string templatePath = "template.docx";
        template.Save(templatePath, SaveFormat.Docx);

        // Load the template for report generation
        Document report = new Document(templatePath);

        // Build the report using the LINQ Reporting engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Save the generated report
        const string outputPath = "report.docx";
        report.Save(outputPath, SaveFormat.Docx);

        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}

// Public data model classes
public class Item
{
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}
