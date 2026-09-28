using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class CategoryCount
{
    public string Category { get; set; } = "";
    public int Count { get; set; }
}

public class ReportModel
{
    public List<CategoryCount> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        const string xmlFileName = "data.xml";
        File.WriteAllText(xmlFileName,
@"<Items>
    <Item><Category>Fruits</Category></Item>
    <Item><Category>Vegetables</Category></Item>
    <Item><Category>Fruits</Category></Item>
    <Item><Category>Dairy</Category></Item>
    <Item><Category>Vegetables</Category></Item>
    <Item><Category>Fruits</Category></Item>
</Items>");

        // Load XML and aggregate by category using GroupBy.
        XDocument xdoc = XDocument.Load(xmlFileName);
        var groups = xdoc.Root!
            .Elements("Item")
            .Select(item => (string?)item.Element("Category") ?? "")
            .GroupBy(cat => cat)
            .Select(g => new CategoryCount { Category = g.Key, Count = g.Count() })
            .ToList();

        // Prepare the model for the report.
        var model = new ReportModel { Groups = groups };

        // Create the template document programmatically.
        const string templateFileName = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Category Report");
        builder.Writeln("<<foreach [g in Groups]>>");
        builder.Writeln("Category: <<[g.Category]>>, Count: <<[g.Count]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templateFileName);

        // Load the template for reporting.
        var template = new Document(templateFileName);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        const string outputFileName = "report.docx";
        template.Save(outputFileName);
    }
}
