using System;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for UTF‑8 and other encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths for the template, output document and XML data source.
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";
        const string xmlPath = "data.xml";

        // -----------------------------------------------------------------
        // 1. Create sample XML data source.
        // -----------------------------------------------------------------
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<order>
    <IncludeHeader>true</IncludeHeader>
    <Title>Sample Report Title</Title>
    <ShowDetails>true</ShowDetails>
    <Details>
        <Detail>First detail line</Detail>
        <Detail>Second detail line</Detail>
        <Detail>Third detail line</Detail>
    </Details>
</order>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Report Begin ===");

        // Conditional header.
        builder.Writeln("<<if [order.IncludeHeader]>>");
        builder.Writeln("Header: <<[order.Title]>>");
        builder.Writeln("<</if>>");

        // Conditional details with a foreach loop.
        builder.Writeln("<<if [order.ShowDetails]>>");
        builder.Writeln("Details:");
        builder.Writeln("<<foreach [d in order.Details.Detail]>>");
        builder.Writeln("- <<[d]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</if>>");

        builder.Writeln("=== Report End ===");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Use the XML file path as the data source.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // Build the report; root object name is "order".
        engine.BuildReport(reportDoc, xmlDataSource, "order");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
