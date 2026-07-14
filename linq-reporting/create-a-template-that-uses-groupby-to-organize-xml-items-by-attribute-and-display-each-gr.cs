using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML loading.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template, XML data and the final report.
        string templatePath = "Template.docx";
        string xmlDataPath = "Data.xml";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("LINQ Reporting – GroupBy Example");
        builder.Writeln();

        // Outer foreach iterates over groups.
        builder.Writeln("<<foreach [g in Groups]>>");
        builder.Writeln("Group: <<[g.Key]>>");
        builder.Writeln();

        // Inner foreach iterates over items within the current group.
        builder.Writeln("<<foreach [i in g.Items]>>");
        builder.Writeln("- <<[i.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before BuildReport).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a sample XML data source.
        // -----------------------------------------------------------------
        string xmlContent =
@"<Items>
    <Item type=""A""><Name>Alpha</Name></Item>
    <Item type=""B""><Name>Beta</Name></Item>
    <Item type=""A""><Name>Gamma</Name></Item>
    <Item type=""C""><Name>Delta</Name></Item>
    <Item type=""B""><Name>Epsilon</Name></Item>
</Items>";

        File.WriteAllText(xmlDataPath, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 3. Load XML, group items by the 'type' attribute and build a model.
        // -----------------------------------------------------------------
        XDocument xDoc = XDocument.Load(xmlDataPath);

        var groups = xDoc.Root!
            .Elements("Item")
            .GroupBy(x => (string?)x.Attribute("type") ?? string.Empty)
            .Select(g => new Group
            {
                Key = g.Key,
                Items = g.Select(i => new Item
                {
                    Name = (string?)i.Element("Name") ?? string.Empty
                }).ToList()
            })
            .ToList();

        ReportModel model = new ReportModel { Groups = groups };

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        // Load the template document.
        Document reportDoc = new Document(templatePath);

        // Build the report. The root object name used in the template tags is "model".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes used by the template.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Group
{
    public string Key { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
}
