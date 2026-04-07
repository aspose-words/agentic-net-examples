using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings used by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample XML data file.
        // -----------------------------------------------------------------
        const string xmlFileName = "Data.xml";
        string xmlContent =
            @"<Orders>
                <Order>
                    <Id>1001</Id>
                    <Customer>John Doe</Customer>
                    <Amount>250.75</Amount>
                </Order>
                <Order>
                    <Id>1002</Id>
                    <Customer>Jane Smith</Customer>
                    <Amount>480.00</Amount>
                </Order>
                <Order>
                    <Id>1003</Id>
                    <Customer>Bob Johnson</Customer>
                    <Amount>125.50</Amount>
                </Order>
              </Orders>";
        File.WriteAllText(xmlFileName, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build the template document programmatically.
        // -----------------------------------------------------------------
        const string templateFileName = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin a foreach block that iterates over the "Order" elements.
        builder.Writeln("<<foreach [order in Orders]>>");

        // Each order will start in its own section.
        // Insert a section break before the first order to ensure a clean start.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Write order details using LINQ Reporting tags.
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.Customer]>>");
        builder.Writeln("Amount: $<<[order.Amount]>>");

        // End of the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templateFileName);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report using the XML data source.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templateFileName);
        var xmlDataSource = new XmlDataSource(xmlFileName);

        // The data source name ("Orders") must match the name used in the foreach tag.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, xmlDataSource, "Orders");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        const string outputFileName = "Report.docx";
        reportDoc.Save(outputFileName);
    }
}
