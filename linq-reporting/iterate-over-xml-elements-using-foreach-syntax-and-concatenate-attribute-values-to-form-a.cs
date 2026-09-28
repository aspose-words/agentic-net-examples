using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML data.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
    <Order Name=""Apple"" Category=""Fruit"" />
    <Order Name=""Carrot"" Category=""Vegetable"" />
    <Order Name=""Banana"" Category=""Fruit"" />
</Orders>";
        string xmlPath = "data.xml";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // Create a Word template with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a paragraph that iterates over the XML elements.
        builder.Writeln("<<foreach [order in Orders]>>");
        // Concatenate attribute values to form a composite string.
        builder.Writeln("<<[order.Name]>> - <<[order.Category]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Load XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "Orders");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
