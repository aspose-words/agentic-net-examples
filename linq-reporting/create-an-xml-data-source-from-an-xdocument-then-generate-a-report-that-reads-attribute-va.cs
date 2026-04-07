using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create sample XML data using XDocument.
        XDocument xmlDoc = new XDocument(
            new XElement("people",
                new XElement("person",
                    new XAttribute("name", "John Doe"),
                    new XAttribute("age", "30")),
                new XElement("person",
                    new XAttribute("name", "Jane Smith"),
                    new XAttribute("age", "25"))
            )
        );

        string xmlPath = Path.Combine(outputDir, "people.xml");
        xmlDoc.Save(xmlPath);

        // 2. Build a template document that contains LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach loop over the collection named "people".
        builder.Writeln("<<foreach [person in people]>>");
        // Output attribute values using the @ prefix. Attribute names must be enclosed in double quotes.
        builder.Writeln("Name: <<[person.@\"name\"]>>");
        builder.Writeln("Age: <<[person.@\"age\"]>>");
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // 3. Load the template document (simulating a separate load step).
        Document loadedTemplate = new Document(templatePath);

        // 4. Create an XmlDataSource from the XML file.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name "people" must match the name used in the template.
        engine.BuildReport(loadedTemplate, dataSource, "people");

        // 6. Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        loadedTemplate.Save(reportPath);
    }
}
