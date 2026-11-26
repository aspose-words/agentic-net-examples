using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class NumberedListReport
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Define file paths.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string dataPath = Path.Combine(outputDir, "data.xml");
        string reportPath = Path.Combine(outputDir, "report.docx");

        // Create sample XML data.
        File.WriteAllText(dataPath,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<Items>
    <Item><Name>Item One</Name></Item>
    <Item><Name>Item Two</Name></Item>
    <Item><Name>Item Three</Name></Item>
</Items>");

        // Build the template document.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Start a numbered list.
        builder.ListFormat.ApplyNumberDefault();

        // Insert the restartNum tag followed by a foreach loop that outputs each item name.
        // The restartNum tag must be placed in the same numbered paragraph as the foreach tag.
        builder.Writeln("<<restartNum>><<foreach [item in Items]>> <<[item.Name]>> <</foreach>>");

        // End the numbered list.
        builder.ListFormat.RemoveNumbers();

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Load XML data source.
        XmlDataSource xmlData = new XmlDataSource(dataPath);

        // Build the report. The data source name must match the root element used in the template tags.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, xmlData, "Items");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}
