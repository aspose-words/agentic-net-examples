using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare sample XML data with boolean flags.
            const string xmlFileName = "ReportData.xml";
            string xmlContent =
                @"<Report>
                    <ShowSection1>true</ShowSection1>
                    <ShowSection2>false</ShowSection2>
                    <Section1Text>Content of the first conditional section.</Section1Text>
                    <Section2Text>Content of the second conditional section.</Section2Text>
                  </Report>";
            File.WriteAllText(xmlFileName, xmlContent);

            // Create a template document programmatically and insert LINQ Reporting tags.
            const string templateFileName = "ReportTemplate.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("=== Sample Report ===");
            // Conditional block for Section 1.
            builder.Writeln("<<if [report.ShowSection1]>>");
            builder.Writeln("<<[report.Section1Text]>>");
            builder.Writeln("<</if>>");
            // Conditional block for Section 2.
            builder.Writeln("<<if [report.ShowSection2]>>");
            builder.Writeln("<<[report.Section2Text]>>");
            builder.Writeln("<</if>>");

            // Save the template to disk.
            templateDoc.Save(templateFileName);

            // Load the template document for reporting.
            Document reportDoc = new Document(templateFileName);

            // Load the XML data source.
            XmlDataSource dataSource = new XmlDataSource(xmlFileName);

            // Build the report using the data source. The root name is "report".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, dataSource, "report");

            // Save the generated report.
            const string outputFileName = "ReportOutput.docx";
            reportDoc.Save(outputFileName);
        }
    }
}
