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
        // Create sample XML data file.
        const string xmlFileName = "data.xml";
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Report>
    <Title>Sample LINQ Reporting PDF/A Report</Title>
    <Items>
        <Item>
            <Name>Apple</Name>
            <Quantity>10</Quantity>
        </Item>
        <Item>
            <Name>Banana</Name>
            <Quantity>20</Quantity>
        </Item>
        <Item>
            <Name>Cherry</Name>
            <Quantity>15</Quantity>
        </Item>
    </Items>
</Report>";
        File.WriteAllText(xmlFileName, xmlContent);

        // Create the template document with LINQ Reporting tags.
        const string templateFileName = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("<<[Report.Title]>>");
        builder.Writeln();

        // Begin foreach loop over items.
        builder.Writeln("<<foreach [item in Report.Items.Item]>>");

        // Table with header and data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Data row (repeated for each item).
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();

        // End table.
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templateFileName);

        // Load the template for report generation.
        Document reportDoc = new Document(templateFileName);

        // Load XML data source.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlFileName);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, xmlDataSource, "Report");

        // Save the report as PDF/A compliant file.
        const string outputPdf = "report.pdf";
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };
        reportDoc.Save(outputPdf, pdfOptions);
    }
}
