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
    public static void Main(string[] args)
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Prepare XML with CDATA sections.
        // -----------------------------------------------------------------
        string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<catalog>
    <product>
        <name><![CDATA[Widget <Pro>]]></name>
        <description><![CDATA[This is a ""special"" product & more.]]></description>
    </product>
    <product>
        <name><![CDATA[Gadget & Co]]></name>
        <description><![CDATA[Another product with <tags> and ""quotes"".]]></description>
    </product>
</catalog>";

        // -----------------------------------------------------------------
        // 2. Load XML into a strongly‑typed model.
        // -----------------------------------------------------------------
        XDocument doc = XDocument.Parse(xmlContent);
        var model = new Catalog
        {
            Products = doc
                .Descendants("product")
                .Select(p => new Product
                {
                    Name = (string)p.Element("name") ?? string.Empty,
                    Description = (string)p.Element("description") ?? string.Empty
                })
                .ToList()
        };

        // -----------------------------------------------------------------
        // 3. Create the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln("<<foreach [p in Products]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Description: <<[p.Description]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Build the report using the model as the data source.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // The root object name used in the template is "data".
        engine.BuildReport(reportDoc, model, "data");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class Catalog
{
    // The property name must match the name used in the template (Products).
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}
