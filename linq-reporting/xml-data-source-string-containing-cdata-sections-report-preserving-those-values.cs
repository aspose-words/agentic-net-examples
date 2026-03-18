using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // XML string that contains CDATA sections.
        // CDATA ensures that characters like <, >, &, " are kept verbatim.
        string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<persons>
    <person>
        <Name><![CDATA[John <Doe>]]></Name>
        <Bio><![CDATA[John loves coding & coffee.]]></Bio>
    </person>
    <person>
        <Name><![CDATA[Jane & Smith]]></Name>
        <Bio><![CDATA[Jane's favorite quote: ""Stay hungry, stay foolish.""]]></Bio>
    </person>
</persons>";

        // Convert the XML string to a memory stream.
        // Using a stream constructor of XmlDataSource preserves CDATA content exactly.
        using (MemoryStream xmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml)))
        {
            // Create the XML data source from the stream.
            XmlDataSource dataSource = new XmlDataSource(xmlStream);

            // Load the template document that contains reporting tags.
            // Example tag in the template: <<foreach [persons.person]>><<[Name]>> - <<[Bio]>>\n<</foreach>>
            Document template = new Document("Template.docx");

            // Build the report. The third argument is the name used in the template to reference the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource, "persons");

            // Save the generated report.
            template.Save("Report.docx");
        }
    }
}
