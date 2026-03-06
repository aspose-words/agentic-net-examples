using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsXmlReport
{
    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains reporting engine tags (e.g. <<[persons.Person.Name]>>)
            string templatePath = @"C:\Docs\Template.docx";

            // Path to the XML file that will be used as the data source
            string xmlDataPath = @"C:\Docs\People.xml";

            // Path where the populated report will be saved
            string outputPath = @"C:\Docs\ReportResult.docx";

            // Load the template document (create -> load)
            Document template = new Document(templatePath);

            // Create an XML data source from the file (create)
            XmlDataSource xmlDataSource = new XmlDataSource(xmlDataPath);

            // Initialize the reporting engine
            ReportingEngine engine = new ReportingEngine();

            // Populate the template with data from the XML source.
            // The second argument is the data source object.
            // The third argument is the name used to reference the data source inside the template.
            // In the template you would use tags like <<[persons.Person.Name]>> where "persons" matches this name.
            engine.BuildReport(template, xmlDataSource, "persons");

            // Save the generated report (save)
            template.Save(outputPath);
        }
    }
}
