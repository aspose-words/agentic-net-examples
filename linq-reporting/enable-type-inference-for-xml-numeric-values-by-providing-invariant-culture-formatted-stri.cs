using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some data sources)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Create XML data source with invariant‑culture formatted numeric strings
        string xmlPath = Path.Combine(outputDir, "Data.xml");
        File.WriteAllText(xmlPath,
            @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
    <Order>
        <Id>1</Id>
        <Amount>1234.56</Amount>
    </Order>
    <Order>
        <Id>2</Id>
        <Amount>7890.12</Amount>
    </Order>
    <Order>
        <Id>3</Id>
        <Amount>345.67</Amount>
    </Order>
</Orders>", Encoding.UTF8);

        // Create a template document programmatically
        string templatePath = Path.Combine(outputDir, "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>   Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        var doc = new Document(templatePath);

        // Create XmlDataSource (no schema file needed)
        var xmlDataSource = new XmlDataSource(xmlPath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "Orders");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);

        Console.WriteLine($"Report generated: {reportPath}");
    }
}
