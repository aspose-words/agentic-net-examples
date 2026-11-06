using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Sample XML data
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Items>
    <Item><Name>First item</Name></Item>
    <Item><Name>Second item</Name></Item>
    <Item><Name>Third item</Name></Item>
</Items>";
        string xmlPath = Path.Combine(outputDir, "data.xml");
        File.WriteAllText(xmlPath, xmlContent);

        // Create template document
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Apply numbered list style to the paragraph that will contain the foreach loop
        builder.ListFormat.ApplyNumberDefault();

        // LINQ Reporting tags for numbered list (restartNum must precede foreach)
        builder.Writeln("<<restartNum>> <<foreach [item in Items]>> <<[item.Name]>> <</foreach>>");

        // Save template (optional, for inspection)
        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // Load XML data source
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // Build report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, xmlData, "Items");

        // Save final report
        string reportPath = Path.Combine(outputDir, "NumberedListReport.docx");
        template.Save(reportPath);

        Console.WriteLine($"Report generated: {reportPath}");
    }
}
