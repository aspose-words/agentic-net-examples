using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare directories.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        // 1. Create sample XML data and save it.
        XDocument xml = new XDocument(
            new XElement("People",
                new XElement("Person", new XAttribute("Name", "John"), new XAttribute("Age", "30")),
                new XElement("Person", new XAttribute("Name", "Jane"), new XAttribute("Age", "25"))
            )
        );
        string xmlPath = Path.Combine(workDir, "data.xml");
        xml.Save(xmlPath);

        // 2. Build a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin a foreach loop over the collection named "persons".
        builder.Writeln("<<foreach [person in persons]>>");
        // Output attribute values.
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        // End the loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template document.
        Document reportDoc = new Document(templatePath);

        // 4. Create an XML data source from the saved XML file.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, dataSource, "persons");

        // 6. Save the generated report.
        string reportPath = Path.Combine(workDir, "report.docx");
        reportDoc.Save(reportPath);
    }
}
