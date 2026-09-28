using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create sample XML data file.
        string xmlPath = Path.Combine(outputDir, "data.xml");
        string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Order>
    <CustomerName>John Doe</CustomerName>
    <Total>123.45</Total>
</Order>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // Create a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("----------------");
        builder.Writeln("Customer: <<[CustomerName]>>");
        builder.Writeln("Total: $<<[Total]>>");
        builder.Writeln("----------------");

        templateDoc.Save(templatePath);

        // Load the template.
        Document reportDoc = new Document(templatePath);

        // Load XML data source.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(reportDoc, xmlDataSource, "Order");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(outputPath);
    }
}
