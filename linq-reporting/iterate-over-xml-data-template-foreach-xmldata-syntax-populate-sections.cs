using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsXmlForeachDemo
{
    class Program
    {
        static void Main()
        {
            // Create a temporary directory for the demo files.
            string tempDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
            Directory.CreateDirectory(tempDir);

            // Create the template document containing the foreach tag.
            string templatePath = Path.Combine(tempDir, "TemplateWithForeach.docx");
            Document templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("<<foreach [in xmlData]>>");
            builder.Writeln("Name: <<[Name]>>");
            builder.Writeln("Age: <<[Age]>>");
            builder.Writeln("<</foreach>>");
            templateDoc.Save(templatePath);

            // XML data to be used as the data source.
            string xmlContent = @"<People>
  <Person>
    <Name>John Doe</Name>
    <Age>30</Age>
  </Person>
  <Person>
    <Name>Jane Smith</Name>
    <Age>25</Age>
  </Person>
</People>";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create an XmlDataSource from the XML string.
            using var xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent));
            XmlDataSource xmlDataSource = new XmlDataSource(xmlStream);

            // Build the report using the data source named "xmlData".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, xmlDataSource, "xmlData");

            // Save the populated document.
            string outputPath = Path.Combine(tempDir, "ReportFromXml.docx");
            doc.Save(outputPath);

            Console.WriteLine($"Report generated at: {outputPath}");
        }
    }
}
