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
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare file paths
        string workDir = Directory.GetCurrentDirectory();
        string dataFile = Path.Combine(workDir, "data.xml");
        string templateFile = Path.Combine(workDir, "template.docx");
        string outputFile = Path.Combine(workDir, "report.docx");

        // 1. Create sample XML data source
        string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Orders>
    <Order>
        <Id>1</Id>
        <CustomerName>John Doe</CustomerName>
        <Total>123.45</Total>
    </Order>
    <Order>
        <Id>2</Id>
        <CustomerName>Jane Smith</CustomerName>
        <Total>678.90</Total>
    </Order>
</Orders>";
        File.WriteAllText(dataFile, xmlContent, Encoding.UTF8);

        // 2. Build a Word template programmatically with LINQ Reporting tags
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Order Report");
        builder.Writeln();

        // Begin foreach over Orders
        builder.Writeln("<<foreach [order in Orders]>>");

        // Start table inside the foreach block
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell(); builder.Writeln("Id");
        builder.InsertCell(); builder.Writeln("Customer");
        builder.InsertCell(); builder.Writeln("Total");
        builder.EndRow();

        // Data row
        builder.InsertCell(); builder.Writeln("<<[order.Id]>>");
        builder.InsertCell(); builder.Writeln("<<[order.CustomerName]>>");
        builder.InsertCell(); builder.Writeln("<<[order.Total]>>");
        builder.EndRow();

        // End table
        builder.EndTable();

        // End foreach
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templateFile);

        // 3. Load the template and generate the report using the XML data source
        Document reportDoc = new Document(templateFile);
        ReportingEngine engine = new ReportingEngine();

        // Build the report: root name matches the XML root element "Orders"
        engine.BuildReport(reportDoc, new XmlDataSource(dataFile), "Orders");

        // 4. Save the generated report
        reportDoc.Save(outputFile);
    }
}
