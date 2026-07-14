using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    public class Program
    {
        public static void Main()
        {
            // Enable code page support for XML parsing (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Create an output folder for generated files.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a sample XML data source with boolean flags.
            // -----------------------------------------------------------------
            string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Report>
    <Title>Quarterly Sales Report</Title>
    <ShowSection1>true</ShowSection1>
    <ShowSection2>false</ShowSection2>
    <Summary>Overall sales increased by 12% compared to the previous quarter.</Summary>
</Report>";
            string xmlPath = Path.Combine(outputDir, "data.xml");
            File.WriteAllText(xmlPath, xmlContent);

            // -----------------------------------------------------------------
            // 2. Build the template document programmatically.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(outputDir, "template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Title (always shown)
            builder.Writeln("<<[model.Title]>>");
            builder.Writeln();

            // Conditional Section 1
            builder.Writeln("<<if [model.ShowSection1]>>");
            builder.Writeln("Section 1: Detailed analysis of product performance.");
            builder.Writeln("<</if>>");
            builder.Writeln();

            // Conditional Section 2
            builder.Writeln("<<if [model.ShowSection2]>>");
            builder.Writeln("Section 2: Forecast for the next quarter.");
            builder.Writeln("<</if>>");
            builder.Writeln();

            // Summary (always shown)
            builder.Writeln("Summary:");
            builder.Writeln("<<[model.Summary]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template document.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 4. Create XmlDataSource from the XML file.
            // -----------------------------------------------------------------
            XmlDataSource dataSource = new XmlDataSource(xmlPath);

            // -----------------------------------------------------------------
            // 5. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root object name used in the template tags is "model".
            engine.BuildReport(reportDoc, dataSource, "model");

            // -----------------------------------------------------------------
            // 6. Save the generated report.
            // -----------------------------------------------------------------
            string resultPath = Path.Combine(outputDir, "report.docx");
            reportDoc.Save(resultPath);

            Console.WriteLine($"Report generated successfully at: {resultPath}");
        }
    }
}
