using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Category { get; set; } = "";
    public decimal Amount { get; set; }
}

public class CategorySummary
{
    public string Category { get; set; } = "";
    public decimal Total { get; set; }
}

public class ReportData
{
    public List<CategorySummary> Summaries { get; set; } = new();
}

class Program
{
    static void Main()
    {
        // Ensure Aspose.Words license is not required for this example.
        // 1. Create sample XML data.
        const string xmlFile = "sample.xml";
        File.WriteAllText(xmlFile,
@"<Items>
    <Item>
        <Category>Food</Category>
        <Amount>10.5</Amount>
    </Item>
    <Item>
        <Category>Food</Category>
        <Amount>20</Amount>
    </Item>
    <Item>
        <Category>Books</Category>
        <Amount>15</Amount>
    </Item>
    <Item>
        <Category>Books</Category>
        <Amount>5</Amount>
    </Item>
    <Item>
        <Category>Electronics</Category>
        <Amount>99.99</Amount>
    </Item>
</Items>");

        // 2. Load XML into objects.
        var items = XDocument.Load(xmlFile)
            .Root!
            .Elements("Item")
            .Select(x => new Item
            {
                Category = (string?)x.Element("Category") ?? "",
                Amount = decimal.Parse((string?)x.Element("Amount") ?? "0")
            })
            .ToList();

        // 3. Summarize by category using LINQ GroupBy.
        var reportData = new ReportData
        {
            Summaries = items
                .GroupBy(i => i.Category)
                .Select(g => new CategorySummary
                {
                    Category = g.Key,
                    Total = g.Sum(i => i.Amount)
                })
                .ToList()
        };

        // 4. Create a template document programmatically.
        const string templateFile = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Category Summary Report");
        builder.Writeln();
        // LINQ Reporting foreach tag.
        builder.Writeln("<<foreach [summary in Summaries]>>");
        builder.Writeln("Category: <<[summary.Category]>>   Total: <<[summary.Total]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templateFile);

        // 5. Load the template and build the report.
        var template = new Document(templateFile);
        var engine = new ReportingEngine();
        engine.BuildReport(template, reportData, "model");

        // 6. Save the final report.
        const string outputFile = "Report.docx";
        template.Save(outputFile);
    }
}
