using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CategoryGroup
{
    public string Category { get; set; } = "";
    public int Count { get; set; }
}

public class ReportModel
{
    public List<CategoryGroup> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create sample XML data file.
        string xmlPath = Path.Combine(outputDir, "data.xml");
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
        List<CategoryGroup> groups = xdoc.Root?
            .Elements("Item")
            .GroupBy(item => (string?)item.Element("Category") ?? "")
            .Select(g => new CategoryGroup
            {
                Category = g.Key,
                Count = g.Count()
            })
            .ToList() ?? new List<CategoryGroup>();

        // Prepare the root model for the report.
        ReportModel model = new() { Groups = groups };

        // Create the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Category Report");
        builder.Writeln("<<foreach [g in Groups]>>");
        builder.Writeln("Category: <<[g.Category]>>, Count: <<[g.Count]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using Aspose.Words LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(reportPath);
    }
}
