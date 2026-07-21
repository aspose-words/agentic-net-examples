using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for XML handling.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template, XML data and the final report.
            const string templatePath = "NumberedListTemplate.docx";
            const string xmlDataPath = "Items.xml";
            const string outputPath = "NumberedListReport.docx";

            // -----------------------------------------------------------------
            // 1. Create a simple XML data file with a collection of items.
            // -----------------------------------------------------------------
            const string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Items>
    <Item><Name>Alpha</Name></Item>
    <Item><Name>Beta</Name></Item>
    <Item><Name>Gamma</Name></Item>
    <Item><Name>Delta</Name></Item>
</Items>";
            File.WriteAllText(xmlDataPath, xmlContent, Encoding.UTF8);

            // -----------------------------------------------------------------
            // 2. Build the template document programmatically.
            //    The template contains a numbered list that iterates over the XML items.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Apply a numbered list style to the paragraph that will hold the foreach block.
            List numberedList = templateDoc.Lists.Add(ListTemplate.NumberDefault);
            builder.ListFormat.List = numberedList;

            // LINQ Reporting tags:
            //   <<foreach [item in Items]>>  – iterate over the collection named "Items".
            //   <<[item.Name]>>              – output the Name property of each item.
            //   <</foreach>>                 – end of the loop.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("<<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk before loading it for the report generation.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and bind the XML data source.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            XmlDataSource dataSource = new XmlDataSource(xmlDataPath);

            // The data source name ("Items") must match the name used in the foreach tag.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, dataSource, "Items");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(outputPath);
        }
    }
}
