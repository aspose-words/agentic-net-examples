using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsXmlReport
{
    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains <<foreach>> tags.
            string templatePath = @"C:\Docs\Template.docx";

            // Path to the XML file that will be used as the data source.
            string xmlPath = @"C:\Docs\Data.xml";

            // Path where the generated report will be saved.
            string outputPath = @"C:\Docs\Report.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create an XML data source from the XML file.
            XmlDataSource dataSource = new XmlDataSource(xmlPath);

            // Build the report. The third argument is the name used in the template
            // to reference the data source (e.g., <<[persons]>><<foreach [in persons]>><<[Name]>>...</foreach>>).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "persons");

            // Save the populated document.
            doc.Save(outputPath);
        }
    }
}
