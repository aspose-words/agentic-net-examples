using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // XML data with CDATA sections. The <items> wrapper is kept, but we enable
        // AlwaysGenerateRootObject so the engine can navigate root -> items -> item.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<root>
    <items>
        <item>
            <Name><![CDATA[Item <One>]]></Name>
            <Description><![CDATA[Description with special characters & < >]]></Description>
        </item>
        <item>
            <Name><![CDATA[Second Item]]></Name>
            <Description><![CDATA[Another description]]></Description>
        </item>
    </items>
</root>";

        // Write the XML string to a memory stream (UTF‑8).
        using MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent));

        // Configure XML loading options to always generate a root object.
        XmlDataLoadOptions loadOptions = new XmlDataLoadOptions { AlwaysGenerateRootObject = true };
        XmlDataSource dataSource = new XmlDataSource(xmlStream, loadOptions);

        // -----------------------------------------------------------------
        // Step 1: Build the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated from XML with CDATA sections");
        builder.Writeln();

        // Begin a foreach block that iterates over each <item> element.
        builder.Writeln("<<foreach [item in root.items.item]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("Description: <<[item.Description]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and generate the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine();
        // Build the report using the XML data source; the root object name is "root".
        engine.BuildReport(reportDoc, dataSource, "root");

        // Save the final report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
