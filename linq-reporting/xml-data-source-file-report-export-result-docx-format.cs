using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportExample
{
    class Program
    {
        static void Main()
        {
            // Create temporary files for the template, XML data, and output.
            string templatePath = Path.Combine(Path.GetTempPath(), "ReportTemplate.docx");
            string xmlDataPath = Path.Combine(Path.GetTempPath(), "ReportData.xml");
            string outputPath = Path.Combine(Path.GetTempPath(), "GeneratedReport.docx");

            // Build a simple Word template containing a reporting tag.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            // The tag format for ReportingEngine is <?Data/Employee/Name?>
            builder.Writeln("Employee Name: <?Data/Employee/Name?>");
            doc.Save(templatePath);

            // Create a minimal XML data source.
            string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Data>
    <Employee>
        <Name>John Doe</Name>
    </Employee>
</Data>";
            File.WriteAllText(xmlDataPath, xmlContent);

            // Load the template document.
            Document templateDoc = new Document(templatePath);

            // Create an XML data source from the file.
            XmlDataSource dataSource = new XmlDataSource(xmlDataPath);

            // Build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(templateDoc, dataSource, "Data");

            // Save the generated report.
            templateDoc.Save(outputPath);

            Console.WriteLine($"Report generated successfully at: {outputPath}");
        }
    }
}
