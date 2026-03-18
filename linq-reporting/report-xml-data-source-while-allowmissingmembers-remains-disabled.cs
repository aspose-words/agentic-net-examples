using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the template document that contains the reporting tags.
        Document template = new Document("Template.docx");

        // Create an XML data source from a file. This will be used by the reporting engine.
        XmlDataSource xmlData = new XmlDataSource("Data.xml");

        // Initialize the reporting engine. Do NOT enable AllowMissingMembers (default is disabled).
        ReportingEngine engine = new ReportingEngine();

        // Build the report. The XML data source is exposed to the template with the name "persons".
        engine.BuildReport(template, xmlData, "persons");

        // Save the populated document.
        template.Save("Report.docx");
    }
}
