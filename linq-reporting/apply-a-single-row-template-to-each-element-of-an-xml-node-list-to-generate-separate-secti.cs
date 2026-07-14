using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingXmlExample
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings that Aspose.Words might need.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample XML data file.
        // -----------------------------------------------------------------
        const string xmlFileName = "data.xml";
        string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Items>
    <Item>
        <Title>First Item</Title>
        <Description>This is the first item's description.</Description>
    </Item>
    <Item>
        <Title>Second Item</Title>
        <Description>This is the second item's description.</Description>
    </Item>
    <Item>
        <Title>Third Item</Title>
        <Description>This is the third item's description.</Description>
    </Item>
</Items>";
        File.WriteAllText(xmlFileName, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Build a template document that contains a foreach tag which will be
        //    applied to each <Item> element.
        // -----------------------------------------------------------------
        const string templateFileName = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Opening foreach tag – iterate over each <Item> element.
        // The root object name is "Items", so we iterate over the collection "Items".
        builder.Writeln("<<foreach [item in Items]>>");

        // Insert a section break so each item starts in a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Content that will be repeated for every <Item>.
        builder.Writeln("Title: <<[item.Title]>>");
        builder.Writeln("Description: <<[item.Description]>>");

        // Closing foreach tag.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templateFileName);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the XML data source.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templateFileName);

        // Load XML data source from the file stream.
        using (FileStream xmlStream = File.OpenRead(xmlFileName))
        {
            XmlDataSource xmlDataSource = new XmlDataSource(xmlStream);

            // Build the report. The root object name must match the top‑level XML element.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, xmlDataSource, "Items");
        }

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        const string outputFileName = "output.docx";
        reportDoc.Save(outputFileName);
    }
}
