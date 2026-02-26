using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the DOTM template that contains LINQ Reporting Engine tags.
            string templatePath = @"C:\Templates\ReportTemplate.dotm";

            // Path to the XML data source file.
            string xmlDataPath = @"C:\Data\People.xml";

            // Path where the generated report will be saved.
            string outputPath = @"C:\Reports\GeneratedReport.docx";

            // Load the DOTM template document.
            Document template = new Document(templatePath);

            // Create an XML data source from the file.
            XmlDataSource dataSource = new XmlDataSource(xmlDataPath);

            // Initialize the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The data source name "persons" must match the name used in the template tags.
            engine.BuildReport(template, dataSource, "persons");

            // Save the populated document.
            template.Save(outputPath);
        }
    }
}
