using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;          // Needed for ListTemplate
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create sample XML data.
        string xmlPath = Path.Combine(workDir, "Data.xml");
        string xmlContent =
            @"<Items>
                <Item>
                    <Index>1</Index>
                    <Name>Apple</Name>
                </Item>
                <Item>
                    <Index>2</Index>
                    <Name>Banana</Name>
                </Item>
                <Item>
                    <Index>3</Index>
                    <Name>Cherry</Name>
                </Item>
              </Items>";
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Build the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Apply a numbered list style to the paragraph that will contain the tags.
        builder.ListFormat.List = templateDoc.Lists.Add(ListTemplate.NumberDefault);

        // Insert the restartNum tag followed by a foreach loop over the XML collection.
        // The foreach iterates over "items" (the root name we will use when building the report).
        builder.Writeln("<<restartNum>><<foreach [item in items]>> <<[item.Name]>> <</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // 4. Create an XmlDataSource from the XML file.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name must match the collection name used in the template ("items").
        engine.BuildReport(reportDoc, dataSource, "items");

        // 6. Save the generated report.
        string reportPath = Path.Combine(workDir, "NumberedListReport.docx");
        reportDoc.Save(reportPath);
    }
}
