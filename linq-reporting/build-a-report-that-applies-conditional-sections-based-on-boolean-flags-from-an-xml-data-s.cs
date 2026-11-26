using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalSections
{
    public class Program
    {
        public static void Main()
        {
            // Ensure the working directory exists.
            string workDir = Directory.GetCurrentDirectory();

            // 1. Create sample XML data source with boolean flags.
            string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Report>
    <Title>Sample Report</Title>
    <ShowSection1>true</ShowSection1>
    <ShowSection2>false</ShowSection2>
</Report>";
            string xmlPath = Path.Combine(workDir, "data.xml");
            File.WriteAllText(xmlPath, xmlContent);

            // 2. Build the template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Title placeholder.
            builder.Writeln("Report Title: <<[data.Title]>>");
            builder.Writeln();

            // Conditional Section 1.
            builder.Writeln("<<if [data.ShowSection1]>>");
            builder.Writeln("=== Section 1 ===");
            builder.Writeln("This content is shown because ShowSection1 is true.");
            builder.Writeln("<</if>>");
            builder.Writeln();

            // Conditional Section 2.
            builder.Writeln("<<if [data.ShowSection2]>>");
            builder.Writeln("=== Section 2 ===");
            builder.Writeln("This content is shown because ShowSection2 is true.");
            builder.Writeln("<</if>>");

            // Save the template to disk.
            string templatePath = Path.Combine(workDir, "Template.docx");
            template.Save(templatePath);

            // 3. Load the template for report generation.
            Document reportDoc = new Document(templatePath);

            // 4. Create the XML data source.
            XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

            // 5. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, xmlDataSource, "data");

            // 6. Save the generated report.
            string reportPath = Path.Combine(workDir, "Report.docx");
            reportDoc.Save(reportPath);

            // Optional: indicate completion (no interactive input).
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
