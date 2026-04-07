using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ConditionalReportExample
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string xmlDataPath = Path.Combine(workDir, "data.xml");
        string outputPath = Path.Combine(workDir, "report_output.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title of the report.
        builder.Writeln("<<[report.Title]>>");
        builder.Writeln();

        // Conditional section – displayed only when ShowDiscount is true.
        builder.Writeln("<<if [report.ShowDiscount]>>");
        builder.Writeln("Discount Applied: <<[report.DiscountAmount]>>");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a simple XML data source file.
        // -----------------------------------------------------------------
        string xmlContent =
@"<Report>
    <Title>Monthly Sales Report</Title>
    <ShowDiscount>true</ShowDiscount>
    <DiscountAmount>15.5</DiscountAmount>
</Report>";
        File.WriteAllText(xmlDataPath, xmlContent);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the XML data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        XmlDataSource dataSource = new XmlDataSource(xmlDataPath);

        // Build the report. The root object name used in the template is "report".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "report");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
