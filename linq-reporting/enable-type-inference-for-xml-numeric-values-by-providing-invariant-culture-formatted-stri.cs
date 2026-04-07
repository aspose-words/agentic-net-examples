using System;
using System.Globalization;
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

        // Define file paths in the current working directory.
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "orders.xml");
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // Create sample XML data with numeric values formatted using invariant culture.
        CreateSampleXml(xmlPath);

        // Build a Word template containing LINQ Reporting tags.
        CreateTemplateDocument(templatePath);

        // Load the template and XML data source.
        Document template = new Document(templatePath);
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "orders");

        // Save the generated report.
        template.Save(outputPath);
    }

    // Generates an XML file with orders. Numeric values are formatted with invariant culture.
    private static void CreateSampleXml(string filePath)
    {
        // Use invariant culture to format numbers (dot as decimal separator).
        string xmlContent =
            $"<?xml version=\"1.0\" encoding=\"UTF-8\"?>{Environment.NewLine}" +
            "<Orders>" + Environment.NewLine +
            "  <Order>" + Environment.NewLine +
            $"    <Id>{1.ToString(CultureInfo.InvariantCulture)}</Id>" + Environment.NewLine +
            "    <CustomerName>John Doe</CustomerName>" + Environment.NewLine +
            $"    <Total>{1234.56.ToString(CultureInfo.InvariantCulture)}</Total>" + Environment.NewLine +
            "  </Order>" + Environment.NewLine +
            "  <Order>" + Environment.NewLine +
            $"    <Id>{2.ToString(CultureInfo.InvariantCulture)}</Id>" + Environment.NewLine +
            "    <CustomerName>Jane Smith</CustomerName>" + Environment.NewLine +
            $"    <Total>{789.00.ToString(CultureInfo.InvariantCulture)}</Total>" + Environment.NewLine +
            "  </Order>" + Environment.NewLine +
            "</Orders>";

        File.WriteAllText(filePath, xmlContent, Encoding.UTF8);
    }

    // Creates a Word document with LINQ Reporting tags to iterate over the XML data.
    private static void CreateTemplateDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Orders Report");
        builder.Writeln();

        // Begin foreach loop over the 'orders' data source.
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Id: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total: <<[order.Total]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}
