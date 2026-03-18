using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a simple template with a conditional section.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        // The IF block will be processed by the ReportingEngine.
        builder.Writeln("<<if [ds.IncludeSection]>>This text appears only when IncludeSection is true.<<endif>>");
        builder.Writeln("This text is always present.");

        // XML data source containing a boolean flag.
        string xmlContent = @"<Root><IncludeSection>true</IncludeSection></Root>";
        using MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent));
        XmlDataSource dataSource = new XmlDataSource(xmlStream);

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Remove paragraphs that become empty after processing.
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report. The data source is referenced in the template as "ds".
        engine.BuildReport(template, dataSource, "ds");

        // Save the generated document.
        template.Save("ConditionalReport.docx");
    }
}
