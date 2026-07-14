using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample XML data with attributes.
        XDocument xDoc = new XDocument(
            new XElement("People",
                new XElement("Person", new XAttribute("Name", "John Doe"), new XAttribute("Age", "30")),
                new XElement("Person", new XAttribute("Name", "Jane Smith"), new XAttribute("Age", "25"))
            )
        );

        // Write the XML to a memory stream.
        using MemoryStream xmlStream = new MemoryStream();
        xDoc.Save(xmlStream);
        xmlStream.Position = 0; // Reset for reading.

        // Create an XML data source from the stream.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlStream);

        // Build a simple Word template programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a foreach tag that iterates over the "Person" elements.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, demonstrates load/save lifecycle).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template back (simulating a separate load step).
        Document doc = new Document(templatePath);

        // Build the report using the XML data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "persons");

        // Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
