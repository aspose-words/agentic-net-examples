using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // File paths.
        string dataFile = Path.Combine(workDir, "data.xml");
        string templateFile = Path.Combine(workDir, "template.docx");
        string outputPdf = Path.Combine(workDir, "Report.pdf");

        // Create sample XML data source.
        File.WriteAllText(dataFile,
@"<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <OrderDate>2023-01-15</OrderDate>
        <Total>199.99</Total>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <OrderDate>2023-02-03</OrderDate>
        <Total>349.50</Total>
    </Order>
</Orders>");

        // Build the template document with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln();

        // Begin foreach loop over Orders.
        builder.Writeln("<<foreach [order in Orders]>>");

        // Create a table for each iteration (header + data row).
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Customer");
        builder.InsertCell();
        builder.Writeln("Date");
        builder.InsertCell();
        builder.Writeln("Total");
        builder.EndRow();

        // Data row bound to the current order.
        builder.InsertCell();
        builder.Writeln("<<[order.CustomerName]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.OrderDate]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Total]>>");
        builder.EndRow();

        // End the table for this iteration.
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templateFile);

        // Load the template for reporting.
        var reportDoc = new Document(templateFile);

        // Load XML data source.
        var xmlData = new XmlDataSource(dataFile);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, xmlData, "Orders");

        // Save as PDF/A-1b.
        var pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };
        reportDoc.Save(outputPdf, pdfOptions);
    }
}
