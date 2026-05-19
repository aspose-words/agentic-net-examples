using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

namespace AsposeWordsLinqReportingDemo
{
    public class Program
    {
        public static void Main()
        {
            // Prepare folders.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create sample XML data.
            string xmlPath = Path.Combine(outputDir, "Data.xml");
            File.WriteAllText(xmlPath,
@"<Items>
    <Item>
        <Index>1</Index>
        <Name>Alpha</Name>
    </Item>
    <Item>
        <Index>2</Index>
        <Name>Beta</Name>
    </Item>
    <Item>
        <Index>3</Index>
        <Name>Gamma</Name>
    </Item>
</Items>");

            // 2. Build the template document programmatically.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Create a numbered list style.
            List numberedList = templateDoc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = numberedList;

            // Insert the LINQ Reporting tags.
            // <<restartNum>> must be placed immediately before the foreach in the same numbered paragraph.
            builder.Writeln("<<restartNum>><<foreach [item in items]>> <<[item.Index]>>. <<[item.Name]>> <</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // 3. Load the template and the XML data source.
            Document doc = new Document(templatePath);
            XmlDataSource dataSource = new XmlDataSource(xmlPath);

            // 4. Build the report.
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple scenario.
            engine.BuildReport(doc, dataSource, "items");

            // 5. Save the generated report.
            string reportPath = Path.Combine(outputDir, "NumberedListReport.docx");
            doc.Save(reportPath);

            // Inform the user (optional, not interactive).
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
