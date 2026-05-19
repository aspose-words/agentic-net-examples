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
        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string dataPath = Path.Combine(workDir, "data.xml");
        string templatePath = Path.Combine(workDir, "template.docx");
        string reportPath = Path.Combine(workDir, "report.docx");

        // 1. Create sample XML data.
        CreateSampleXml(dataPath);

        // 2. Load XML and aggregate by Category using LINQ GroupBy.
        ReportModel model = BuildReportModelFromXml(dataPath);

        // 3. Create a LINQ Reporting template programmatically.
        CreateTemplateDocument(templatePath);

        // 4. Load the template and build the report.
        Document templateDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(templateDoc, model, "model");

        // 5. Save the generated report.
        templateDoc.Save(reportPath);
    }

    // Creates a simple XML file with a list of items, each having a Category element.
    private static void CreateSampleXml(string filePath)
    {
        XDocument doc = new XDocument(
            new XElement("Items",
                new XElement("Item", new XElement("Category", "Fruit")),
                new XElement("Item", new XElement("Category", "Fruit")),
                new XElement("Item", new XElement("Category", "Vegetable")),
                new XElement("Item", new XElement("Category", "Fruit")),
                new XElement("Item", new XElement("Category", "Grain")),
                new XElement("Item", new XElement("Category", "Vegetable"))
            )
        );
        doc.Save(filePath);
    }

    // Loads the XML file, groups items by Category, and builds the model for the report.
    private static ReportModel BuildReportModelFromXml(string xmlPath)
    {
        XDocument xdoc = XDocument.Load(xmlPath);
        var groups = xdoc.Root!
            .Elements("Item")
            .GroupBy(item => (string?)item.Element("Category") ?? string.Empty)
            .Select(g => new GroupInfo
            {
                Category = g.Key,
                Count = g.Count()
            })
            .ToList();

        return new ReportModel { Groups = groups };
    }

    // Generates a Word document containing LINQ Reporting tags.
    private static void CreateTemplateDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Items Count by Category");
        builder.Writeln(""); // empty line

        // Begin foreach loop over Groups collection.
        builder.Writeln("<<foreach [group in Groups]>>");
        // Output each group's Category and Count.
        builder.Writeln("Category: <<[group.Category]>>, Total: <<[group.Count]>>");
        // End foreach loop.
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// Root model passed to the reporting engine.
public class ReportModel
{
    public List<GroupInfo> Groups { get; set; } = new();
}

// Represents a single aggregated group.
public class GroupInfo
{
    public string Category { get; set; } = string.Empty;
    public int Count { get; set; }
}
