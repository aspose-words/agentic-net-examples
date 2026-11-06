using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML data.
        const string xmlPath = "data.xml";
        File.WriteAllText(xmlPath, @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Report>
  <Title>Sample LINQ Reporting</Title>
  <ShowSummary>true</ShowSummary>
  <ShowDetails>true</ShowDetails>
  <Items>
    <Item>
      <Name>Item 1</Name>
      <Value>10</Value>
    </Item>
    <Item>
      <Name>Item 2</Name>
      <Value>20</Value>
    </Item>
    <Item>
      <Name>Item 3</Name>
      <Value>30</Value>
    </Item>
  </Items>
</Report>");

        // Create the template document programmatically.
        const string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("<<[Report.Title]>>");
        builder.Writeln();

        // Conditional Summary Section
        builder.Writeln("<<if [Report.ShowSummary]>>");
        builder.Writeln("=== Summary Section ===");
        builder.Writeln("This is a summary that appears only when ShowSummary is true.");
        builder.Writeln("<</if>>");
        builder.Writeln();

        // Conditional Details Section with a foreach loop
        builder.Writeln("<<if [Report.ShowDetails]>>");
        builder.Writeln("=== Details Section ===");
        builder.Writeln("<<foreach [item in Report.Items.Item]>>");
        builder.Writeln("- <<[item.Name]>> : <<[item.Value]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</if>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Load XML data source.
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, xmlData, "Report");

        // Save the generated report.
        const string outputPath = "output.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}
