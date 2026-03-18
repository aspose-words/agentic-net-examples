using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Build an XML document in memory using XDocument.
        XDocument xDoc = new XDocument(
            new XElement("People",
                new XElement("Person",
                    new XAttribute("Id", 1),
                    new XAttribute("Name", "John"),
                    new XAttribute("Age", 30)),
                new XElement("Person",
                    new XAttribute("Id", 2),
                    new XAttribute("Name", "Jane"),
                    new XAttribute("Age", 28))
            )
        );

        // Write the XDocument to a MemoryStream so it can be consumed by XmlDataSource.
        using (MemoryStream xmlStream = new MemoryStream())
        {
            xDoc.Save(xmlStream);
            xmlStream.Position = 0; // Reset stream position for reading.

            // Create an XmlDataSource from the XML stream.
            XmlDataSource dataSource = new XmlDataSource(xmlStream);

            // Load the Word template that contains reporting tags.
            // Example tag in the template: <<[People.Person.@Id]>>
            Document template = new Document("Template.docx");

            // Build the report. The third argument is the name of the root element
            // that will be used as the data source name inside the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource, "People");

            // Save the populated document.
            template.Save("Report.docx");
        }
    }
}
