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
        string xmlPath = "sample.xml";
        File.WriteAllText(xmlPath,
            @"<Items>
                <Item><Category>Fruit</Category></Item>
                <Item><Category>Vegetable</Category></Item>
                <Item><Category>Fruit</Category></Item>
                <Item><Category>Dairy</Category></Item>
                <Item><Category>Vegetable</Category></Item>
                <Item><Category>Fruit</Category></Item>
              </Items>");

        // Load XML and aggregate by category using LINQ GroupBy.
        XDocument xdoc = XDocument.Load(xmlPath);
        List<CategoryGroup> groups = xdoc.Root!
            .Elements("Item")
            .GroupBy(item => (string?)item.Element("Category") ?? string.Empty)
            .Select(g => new CategoryGroup
            {
                Category = g.Key,
                Count = g.Count()
            })
            .ToList();

        // Wrap the grouped data for the reporting engine.
        ReportModel model = new ReportModel { Groups = groups };

        // Create the template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Category Report");
        builder.Writeln("<<foreach [group in Groups]>>");
        builder.Writeln("Category: <<[group.Category]>>, Total: <<[group.Count]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template and build the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "ReportOutput.docx";
        reportDoc.Save(outputPath);
    }
}

// Model representing a single category group.
public class CategoryGroup
{
    public string Category { get; set; } = string.Empty;
    public int Count { get; set; }
}

// Wrapper model passed to the reporting engine.
public class ReportModel
{
    public List<CategoryGroup> Groups { get; set; } = new();
}
