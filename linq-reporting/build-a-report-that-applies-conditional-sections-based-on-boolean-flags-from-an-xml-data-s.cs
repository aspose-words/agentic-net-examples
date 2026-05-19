using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working folder and file paths.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string xmlPath = Path.Combine(workDir, "data.xml");
        string outputPath = Path.Combine(workDir, "report.docx");

        // 1. Create sample XML data source.
        string xmlContent =
@"<Report>
    <Item>
        <Name>Item 1</Name>
        <ShowDetails>true</ShowDetails>
        <Details>Details for Item 1</Details>
    </Item>
    <Item>
        <Name>Item 2</Name>
        <ShowDetails>false</ShowDetails>
        <Details>Details for Item 2</Details>
    </Item>
    <Item>
        <Name>Item 3</Name>
        <ShowDetails>true</ShowDetails>
        <Details>Details for Item 3</Details>
    </Item>
</Report>";
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Build the template document with correct LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The XML root element is treated as a collection, so iterate directly over the root.
        builder.Writeln("<<foreach [item in report]>>");
        builder.Writeln("Item: <<[item.Name]>>");
        builder.Writeln("<<if [item.ShowDetails]>>Details: <<[item.Details]>> <</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Load XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // 5. Build the report using the data source name "report".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, dataSource, "report");

        // 6. Save the generated report.
        reportDoc.Save(outputPath);
    }
}
