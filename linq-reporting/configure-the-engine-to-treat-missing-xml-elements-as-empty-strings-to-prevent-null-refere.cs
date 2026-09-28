using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingMissingElementsExample
{
    public static void Main()
    {
        // Create a folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Order Report");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        // The CustomerName element may be missing in the XML data.
        builder.Writeln("Customer Name: <<[order.CustomerName]>>");
        builder.Writeln("End of Report");

        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // -------------------------------------------------
        // 2. Prepare XML data source.
        //    To treat a missing element as an empty string we add an empty <CustomerName> element.
        // -------------------------------------------------
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<order>
    <OrderId>12345</OrderId>
    <CustomerName></CustomerName> <!-- Empty element to avoid missing-field errors -->
</order>";
        using MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent));
        xmlStream.Position = 0;
        XmlDataSource xmlDataSource = new XmlDataSource(xmlStream);

        // -------------------------------------------------
        // 3. Configure the ReportingEngine (default options are sufficient).
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // -------------------------------------------------
        // 4. Build the report.
        // -------------------------------------------------
        Document report = new Document(templatePath);
        engine.BuildReport(report, xmlDataSource, "order");

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        string reportPath = Path.Combine(outputDir, "Report.docx");
        report.Save(reportPath);
    }
}
