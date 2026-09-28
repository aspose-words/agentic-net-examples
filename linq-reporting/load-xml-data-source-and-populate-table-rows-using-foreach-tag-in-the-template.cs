using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(workDir);

        // Create sample XML data source.
        string xmlPath = Path.Combine(workDir, "data.xml");
        File.WriteAllText(xmlPath,
            @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Orders>
    <Order>
        <Id>1001</Id>
        <Customer>John Doe</Customer>
    </Order>
    <Order>
        <Id>1002</Id>
        <Customer>Jane Smith</Customer>
    </Order>
    <Order>
        <Id>1003</Id>
        <Customer>Bob Johnson</Customer>
    </Order>
</Orders>", Encoding.UTF8);

        // Create the template document programmatically.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Orders Report");
        builder.Writeln();

        // Header table (appears once).
        Table headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Order ID");
        builder.InsertCell();
        builder.Writeln("Customer");
        builder.EndRow();
        builder.EndTable();

        // Begin foreach loop over Orders.
        builder.Writeln("<<foreach [order in Orders]>>");

        // Data rows table (repeated for each order).
        Table dataTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[order.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Customer]>>");
        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using the XML data source.
        ReportingEngine engine = new ReportingEngine();
        XmlDataSource xmlData = new XmlDataSource(xmlPath);
        engine.BuildReport(reportDoc, xmlData, "Orders");

        // Save the generated report.
        string reportPath = Path.Combine(workDir, "report.docx");
        reportDoc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
