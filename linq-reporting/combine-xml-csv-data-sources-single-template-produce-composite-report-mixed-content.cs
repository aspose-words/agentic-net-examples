using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace CompositeReportExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template document that contains tags for both XML and CSV data.
            string templatePath = @"C:\Templates\CompositeReportTemplate.docx";

            // Paths to the XML and CSV data files.
            string xmlDataPath = @"C:\Data\People.xml";
            string csvDataPath = @"C:\Data\Sales.csv";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create data source objects.
            XmlDataSource xmlSource = new XmlDataSource(xmlDataPath);
            CsvDataSource csvSource = new CsvDataSource(csvDataPath);

            // Build the report using both data sources.
            // The first source name can be empty if you reference its members directly,
            // but we give explicit names to use them in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(
                doc,
                new object[] { xmlSource, csvSource },
                new string[] { "xmlData", "csvData" });

            // Save the generated report.
            string outputPath = @"C:\Output\CompositeReport.docx";
            doc.Save(outputPath);
        }
    }
}
