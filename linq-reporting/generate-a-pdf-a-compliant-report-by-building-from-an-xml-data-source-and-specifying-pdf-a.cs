using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // Needed for Table class

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML parsing (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare working folder.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create sample XML data source.
        string xmlPath = Path.Combine(workDir, "orders.xml");
        File.WriteAllText(xmlPath, @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Orders>
    <Order>
        <Id>1</Id>
        <CustomerName>John Doe</CustomerName>
        <Amount>123.45</Amount>
    </Order>
    <Order>
        <Id>2</Id>
        <CustomerName>Jane Smith</CustomerName>
        <Amount>678.90</Amount>
    </Order>
</Orders>");

        // 2. Build the template document programmatically.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln("Generated on: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
        builder.Writeln(); // empty line

        // Begin foreach loop over the XML data source named "orders".
        builder.Writeln("<<foreach [order in orders]>>");

        // Create a table with a header row.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("Id");
        builder.InsertCell();
        builder.Write("Customer");
        builder.InsertCell();
        builder.Write("Amount");
        builder.EndRow();

        // Data row – values are filled by the reporting engine.
        builder.InsertCell();
        builder.Write("<<[order.Id]>>");
        builder.InsertCell();
        builder.Write("<<[order.CustomerName]>>");
        builder.InsertCell();
        builder.Write("<<[order.Amount]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template document.
        Document reportDoc = new Document(templatePath);

        // 4. Create the XML data source.
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // 5. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(reportDoc, xmlData, "orders");

        // 6. Save the report as PDF/A‑1b.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };
        string outputPdf = Path.Combine(workDir, "OrdersReport.pdf");
        reportDoc.Save(outputPdf, pdfOptions);
    }
}
