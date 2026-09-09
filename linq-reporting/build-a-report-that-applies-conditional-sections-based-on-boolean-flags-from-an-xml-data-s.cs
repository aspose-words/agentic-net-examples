using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for older encodings (required by Aspose.Words on .NET Core)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create a simple XML data source file with a boolean flag.
        // -----------------------------------------------------------------
        const string xmlPath = "reportData.xml";
        string xmlContent =
@"<Report>
    <Title>Conditional Sections Example</Title>
    <ShowSection>true</ShowSection>
</Report>";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a Word template programmatically and insert LINQ Reporting tags.
        // -----------------------------------------------------------------
        const string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title placeholder
        builder.Writeln("Report Title: <<[report.Title]>>");
        builder.Writeln();

        // Conditional section – displayed only when ShowSection is true
        builder.Writeln("<<if [report.ShowSection]>>");
        builder.Writeln(">>> This paragraph appears because ShowSection is TRUE.");
        builder.Writeln("<</if>>");

        // Optional else‑like block – displayed when ShowSection is false
        builder.Writeln("<<if [report.ShowSection == false]>>");
        builder.Writeln(">>> This paragraph appears because ShowSection is FALSE.");
        builder.Writeln("<</if>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and the XML data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The root object name used in the template tags is "report"
        engine.BuildReport(doc, dataSource, "report");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}
