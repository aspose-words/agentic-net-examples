using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create sample XML data source.
        string xmlPath = Path.Combine(workDir, "ReportData.xml");
        string xmlContent =
@"<Report>
    <ShowSection1>true</ShowSection1>
    <ShowSection2>false</ShowSection2>
    <Section1Text>Details for section 1.</Section1Text>
    <Section2Text>Details for section 2.</Section2Text>
</Report>";
        File.WriteAllText(xmlPath, xmlContent);

        // Create the template document with conditional sections.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report Title");
        builder.Writeln("<<if [report.ShowSection1]>>");
        builder.Writeln("Section 1:");
        builder.Writeln("<<[report.Section1Text]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<<if [report.ShowSection2]>>");
        builder.Writeln("Section 2:");
        builder.Writeln("<<[report.Section2Text]>>");
        builder.Writeln("<</if>>");

        templateDoc.Save(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Load XML data source.
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(doc, xmlData, "report");

        // Save the generated report.
        string outputPath = Path.Combine(workDir, "ReportResult.docx");
        doc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated at: {outputPath}");
    }
}
