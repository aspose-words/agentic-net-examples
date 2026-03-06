using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class MailMergeXmlExample
{
    static void Main()
    {
        // Path to the DOCX template that contains merge fields (e.g. <<[Customer/Name]>>).
        string templatePath = "Template.docx";

        // Path to the XML file that holds the data for the merge.
        string xmlDataPath = "Data.xml";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Create a ReportingEngine instance – it will process the template.
        ReportingEngine engine = new ReportingEngine();

        // Load the XML data source.
        XmlDataSource xmlData = new XmlDataSource(xmlDataPath);

        // Build the report by merging the XML data into the template.
        engine.BuildReport(doc, xmlData);

        // Save the merged document.
        doc.Save("MergedOutput.docx");
    }
}
