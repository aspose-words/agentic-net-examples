using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Newtonsoft.Json; // Included as required by category

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create sample XML data source.
        string xmlPath = Path.Combine(outputDir, "data.xml");
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Report>
  <Title>Sample Report</Title>
  <Date>2023-01-01</Date>
  <Items>
    <Item>
      <Name>Item 1</Name>
      <Quantity>2</Quantity>
      <Price>10.5</Price>
    </Item>
    <Item>
      <Name>Item 2</Name>
      <Quantity>5</Quantity>
      <Price>7.25</Price>
    </Item>
  </Items>
</Report>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // Create the Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report Title: <<[Report.Title]>>");
        builder.Writeln("Date: <<[Report.Date]>>");
        builder.Writeln("");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in Report.Items.Item]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>> x $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for processing.
        Document doc = new Document(templatePath);

        // Load XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "Report");

        // Save the generated document as PDF/A-1b.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };
        string pdfPath = Path.Combine(outputDir, "report.pdf");
        doc.Save(pdfPath, pdfOptions);
    }
}
