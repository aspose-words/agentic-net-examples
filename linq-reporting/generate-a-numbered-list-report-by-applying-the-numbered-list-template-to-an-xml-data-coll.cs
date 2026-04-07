using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create sample XML data.
        string xmlPath = Path.Combine(outputDir, "Data.xml");
        File.WriteAllText(xmlPath,
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
</Items>");

        // 2. Build the template document programmatically.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Create a numbered list style.
        List numberedList = templateDoc.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = numberedList;

        // Insert a numbered paragraph that will be repeated for each XML item.
        // <<restartNum>> must be placed immediately before the <<foreach>> tag in the same paragraph.
        builder.Writeln("<<restartNum>><<foreach [item in Items]>><<[item.Index]>>. <<[item.Name]>> <</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template and the XML data source.
        Document doc = new Document(templatePath);
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // 4. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The root name in the template is "Items", so we pass it as the data source name.
        engine.BuildReport(doc, xmlData, "Items");

        // 5. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);

        // Inform that the process completed (no interactive input required).
        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
