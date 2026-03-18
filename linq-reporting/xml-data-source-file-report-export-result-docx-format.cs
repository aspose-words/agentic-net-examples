using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template Word document that contains the report tags.
            string templatePath = @"C:\Data\ReportTemplate.docx";

            // Path to the XML file that will be used as the data source.
            string xmlDataPath = @"C:\Data\ReportData.xml";

            // Path where the generated report will be saved in DOCX format.
            string outputPath = @"C:\Data\GeneratedReport.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create an XML data source from the file using the default loading options.
            XmlDataSource dataSource = new XmlDataSource(xmlDataPath);

            // Build the report. The third parameter is the name used in the template to reference the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "Data");

            // Save the populated document as DOCX.
            doc.Save(outputPath);
        }
    }
}
