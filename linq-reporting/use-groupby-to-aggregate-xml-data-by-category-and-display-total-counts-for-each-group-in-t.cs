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
    public List<CategoryCount> Categories { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        const string xmlPath = "Products.xml";
        File.WriteAllText(xmlPath,
            @"<Products>
                <Product><Category>Books</Category></Product>
                <Product><Category>Electronics</Category></Product>
                <Product><Category>Books</Category></Product>
                <Product><Category>Clothing</Category></Product>
                <Product><Category>Electronics</Category></Product>
                <Product><Category>Books</Category></Product>
              </Products>");

        // Load XML and aggregate counts per category using LINQ GroupBy.
        XDocument xdoc = XDocument.Load(xmlPath);
        var grouped = xdoc.Root!
            .Elements("Product")
            .GroupBy(p => (string?)p.Element("Category") ?? "")
            .Select(g => new CategoryCount { Category = g.Key, Count = g.Count() })
            .ToList();

        // Wrap the aggregated data in a model object.
        var model = new ReportModel { Categories = grouped };

        // Create a template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a heading.
        builder.Writeln("Category Summary");
        builder.Writeln();

        // Insert LINQ Reporting tags to iterate over the categories.
        builder.Writeln("<<foreach [c in Categories]>>");
        builder.Writeln("Category: <<[c.Category]>> - Total: <<[c.Count]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, demonstrates load/save lifecycle).
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template back (simulating a real scenario where template is a file).
        var loadedTemplate = new Document(templatePath);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the final report.
        const string reportPath = "Report.docx";
        loadedTemplate.Save(reportPath);
    }
}
