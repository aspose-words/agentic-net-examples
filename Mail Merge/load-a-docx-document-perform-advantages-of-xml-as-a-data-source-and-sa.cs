using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the input DOCX template, the XML data source and the output XPS file.
        string docsPath = @"C:\Docs\";
        string templateFile = Path.Combine(docsPath, "Template.docx");
        string xmlDataFile = Path.Combine(docsPath, "Data.xml");
        string outputFile = Path.Combine(docsPath, "Result.xps");

        // Load the DOCX document.
        Document doc = new Document(templateFile);

        // Create an XML data source from the XML file.
        XmlDataSource xmlSource = new XmlDataSource(xmlDataFile);

        // Build the report using the XML data source.
        // The third argument is the root element name in the XML that contains the data.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, xmlSource, "root");

        // Prepare XPS save options (optional – can be customized here).
        XpsSaveOptions xpsOptions = new XpsSaveOptions();

        // Save the populated document as XPS.
        doc.Save(outputFile, xpsOptions);
    }
}
