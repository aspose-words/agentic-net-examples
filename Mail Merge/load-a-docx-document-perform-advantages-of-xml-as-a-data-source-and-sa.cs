using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the input DOCX template, the XML data source and the output XPS file.
        string docPath = @"C:\Docs\Template.docx";
        string xmlPath = @"C:\Docs\Data.xml";
        string outPath = @"C:\Docs\Result.xps";

        // Load the DOCX document.
        Document doc = new Document(docPath);

        // Create an XML data source from the specified file.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report using the XML data source.
        // ReportingEngine is an instance class; create an instance before calling BuildReport.
        ReportingEngine engine = new ReportingEngine();
        // The third argument is the name of the root element in the XML that contains the data.
        engine.BuildReport(doc, dataSource, "root");

        // Prepare XPS save options (optional – can be customized here).
        XpsSaveOptions saveOptions = new XpsSaveOptions();

        // Save the populated document as XPS.
        doc.Save(outPath, saveOptions);
    }
}
