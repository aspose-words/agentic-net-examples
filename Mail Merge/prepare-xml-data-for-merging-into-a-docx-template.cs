using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsXmlMergeExample
{
    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains reporting engine tags, e.g. <<[persons.Person.Name]>>
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document (lifecycle rule: load)
            Document template = new Document(templatePath);

            // Prepare XML data as a string. The root element name ("Persons") can be any name;
            // the reporting engine will treat it as a collection if it contains a list of identical child elements.
            string xmlContent = @"
                <Persons>
                    <Person>
                        <Name>John Doe</Name>
                        <Age>30</Age>
                        <Email>john.doe@example.com</Email>
                    </Person>
                    <Person>
                        <Name>Jane Smith</Name>
                        <Age>28</Age>
                        <Email>jane.smith@example.com</Email>
                    </Person>
                </Persons>";

            // Convert the XML string to a memory stream (lifecycle rule: create)
            using (MemoryStream xmlStream = new MemoryStream())
            using (StreamWriter writer = new StreamWriter(xmlStream))
            {
                writer.Write(xmlContent);
                writer.Flush();
                xmlStream.Position = 0; // Reset stream position for reading

                // Create an XmlDataSource from the stream (lifecycle rule: create)
                XmlDataSource xmlDataSource = new XmlDataSource(xmlStream);

                // Build the report using the reporting engine.
                // The third parameter ("persons") is the name used in the template to reference the data source.
                ReportingEngine engine = new ReportingEngine();
                bool success = engine.BuildReport(template, xmlDataSource, "persons");

                if (!success)
                {
                    Console.WriteLine("Report building failed. Check template syntax and data.");
                    return;
                }

                // Save the populated document (lifecycle rule: save)
                string outputPath = @"C:\Output\ReportResult.docx";
                template.Save(outputPath);
                Console.WriteLine($"Report generated successfully at: {outputPath}");
            }
        }
    }
}
