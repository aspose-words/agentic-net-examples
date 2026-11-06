using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML data.
        string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Products>
  <Product>
    <Name>Apple</Name>
    <Price>1.20</Price>
  </Product>
  <Product>
    <Name>Banana</Name>
    <Price>0.80</Price>
  </Product>
  <Product>
    <Name>Cherry</Name>
    <Price>2.50</Price>
  </Product>
</Products>";
        string xmlPath = "data.xml";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // Load XML data source.
        var xmlData = new XmlDataSource(xmlPath);

        // Create a template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln();

        // Iterate over the collection of Product nodes.
        builder.Writeln("<<foreach [p in xmlData]>>");
        builder.Writeln("Product: <<[p.Name]>>");
        builder.Writeln("Price: $<<[p.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, just for demonstration).
        string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using the XML data source.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, xmlData, "xmlData");

        // Save the generated report.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
