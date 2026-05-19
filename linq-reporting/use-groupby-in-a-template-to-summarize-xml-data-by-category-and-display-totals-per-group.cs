using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Category { get; set; } = "";
    public string Name { get; set; } = "";
    public double Price { get; set; }
}

public class CategorySummary
{
    public string Category { get; set; } = "";
    public double Total { get; set; }
    public List<Product> Items { get; set; } = new();
}

public class ReportModel
{
    public List<CategorySummary> Categories { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        string xmlPath = "sample.xml";
        File.WriteAllText(xmlPath,
@"<Products>
    <Product>
        <Category>Fruit</Category>
        <Name>Apple</Name>
        <Price>1.2</Price>
    </Product>
    <Product>
        <Category>Fruit</Category>
        <Name>Banana</Name>
        <Price>0.8</Price>
    </Product>
    <Product>
        <Category>Vegetable</Category>
        <Name>Carrot</Name>
        <Price>0.5</Price>
    </Product>
    <Product>
        <Category>Vegetable</Category>
        <Name>Broccoli</Name>
        <Price>1.1</Price>
    </Product>
</Products>");

        // Load XML and build the data model with grouping.
        XDocument xdoc = XDocument.Load(xmlPath);
        var products = xdoc.Root!
            .Elements("Product")
            .Select(p => new Product
            {
                Category = (string)p.Element("Category")!,
                Name = (string)p.Element("Name")!,
                Price = (double)p.Element("Price")!
            })
            .ToList();

        var model = new ReportModel
        {
            Categories = products
                .GroupBy(p => p.Category)
                .Select(g => new CategorySummary
                {
                    Category = g.Key,
                    Total = g.Sum(p => p.Price),
                    Items = g.ToList()
                })
                .ToList()
        };

        // Create the LINQ Reporting template.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.Writeln("Category: <<[cat.Category]>>   Total Price: <<[cat.Total]>>");
        builder.Writeln("Products:");
        builder.Writeln("<<foreach [p in cat.Items]>>");
        builder.Writeln("- <<[p.Name]>> : $<<[p.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
