using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Enable reflection optimization for the ReportingEngine.
        // This improves performance when the engine invokes custom type members via reflection.
        ReportingEngine.UseReflectionOptimization = true;

        // Create a new instance of the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Register a known external type (e.g., System.DateTime) so its static members can be accessed
        // directly from the report template.
        engine.KnownTypes.Add(typeof(DateTime));

        // Load the template document that contains the report placeholders.
        Document template = new Document("Template.docx");

        // Load a large XML data set. The XmlDataSource reads the XML file and makes it available to the engine.
        XmlDataSource xmlData = new XmlDataSource("LargeData.xml");

        // Build the report. The third argument is the name of the top‑level XML element that
        // serves as the root for the data source (adjust as needed for your XML structure).
        engine.BuildReport(template, xmlData, "root");

        // Save the generated report.
        template.Save("ReportOutput.docx");
    }
}
