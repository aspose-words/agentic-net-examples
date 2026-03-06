using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the template document that contains the reporting tags.
        string templatePath = @"MyDir\ReportingTemplate.docx";

        // Path to the prepared XML file that holds the data.
        string xmlPath = @"MyDir\Data.xml";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Create an XmlDataSource from the XML file using the default loading options.
        // The constructor takes the path to the XML file.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // Build the report. The third argument is the name of the root element
        // in the XML that will be referenced in the template (e.g., <<[persons.Name]>>).
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "persons");

        // Save the populated document.
        doc.Save(@"ArtifactsDir\ReportFromXml.docx");
    }
}
