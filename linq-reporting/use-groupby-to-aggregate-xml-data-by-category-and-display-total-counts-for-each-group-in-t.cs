using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directories.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        // 1. Create sample XML data file.
        string xmlPath = Path.Combine(workDir, "data.xml");
        File.WriteAllText(xmlPath,
@"<Items>
    <Item><Category>Fruit</Category></Item>
    <Item><Category>Vegetable</Category></Item>
    <Item><Category>Fruit</Category></Item>
    <Item><Category>Dairy</Category></Item>
    <Item><Category>Vegetable</Category></Item>
    <Item><Category>Fruit</Category></Item>
</Items>");

        // 2. Load XML and aggregate by category using LINQ GroupBy.
        XDocument xdoc = XDocument.Load(xmlPath);
        List<CategoryGroup> groups = xdoc.Root!
            .Elements("Item")
            .GroupBy(item => (string?)item.Element("Category") ?? "")
            .Select(g => new CategoryGroup
            {
                Category = g.Key,
                Count = g.Count()
            })
            .ToList();

        // 3. Create the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Category Report");
        builder.Writeln("==============");
        builder.Writeln("<<foreach [g in Groups]>>");
        builder.Writeln("Category: <<[g.Category]>> - Total Items: <<[g.Count]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // 4. Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // 5. Prepare the root model object.
        ReportModel model = new ReportModel { Groups = groups };

        // 6. Build the report using Aspose.Words ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 7. Save the generated report.
        string reportPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(reportPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
