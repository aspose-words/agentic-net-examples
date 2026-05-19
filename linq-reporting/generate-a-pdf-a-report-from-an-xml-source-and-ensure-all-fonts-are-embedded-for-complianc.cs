using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Tables;   // Needed for Table type

public class PdfAReportGenerator
{
    public static void Main()
    {
        // File paths for the example.
        string dataFile = "data.xml";
        string templateFile = "Template.docx";
        string outputFile = "Report.pdf";

        // 1. Create a simple XML data source.
        string xmlContent =
@"<Report>
    <Title>Sample Report</Title>
    <Date>2023-01-01</Date>
    <Items>
        <Item>
            <Name>Item 1</Name>
            <Quantity>10</Quantity>
        </Item>
        <Item>
            <Name>Item 2</Name>
            <Quantity>5</Quantity>
        </Item>
    </Items>
</Report>";
        File.WriteAllText(dataFile, xmlContent);

        // 2. Build the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<[report.Title]>>");
        builder.Writeln("Date: <<[report.Date]>>");
        builder.Writeln();

        // Begin foreach loop over the Item elements.
        builder.Writeln("<<foreach [item in report.Items.Item]>>");

        // Create a table for each item.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();

        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templateFile);

        // 3. Load the template document.
        Document reportDoc = new Document(templateFile);

        // 4. Create an XmlDataSource from the XML file.
        XmlDataSource xmlData = new XmlDataSource(dataFile);

        // 5. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, xmlData, "report");

        // 6. Save the report as PDF/A with all fonts embedded.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // PDF/A-1b compliance.
            Compliance = PdfCompliance.PdfA1b,
            // Ensure all fonts used in the document are embedded.
            EmbedFullFonts = true
            // FontEmbeddingMode property is optional; EmbedFullFonts is sufficient for most versions.
        };

        reportDoc.Save(outputFile, pdfOptions);
    }
}
