using System;
using System.Globalization;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare directories.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // 1. Create sample XML data with numeric values formatted using invariant culture.
        string xmlPath = Path.Combine(workDir, "people.xml");
        CreateSampleXml(xmlPath);

        // 2. Create a Word template containing LINQ Reporting tags.
        string templatePath = Path.Combine(workDir, "template.docx");
        CreateTemplateDocument(templatePath);

        // 3. Load the template document.
        Document templateDoc = new Document(templatePath);

        // 4. Load the XML data source.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // 5. Build the report.
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this example.
        engine.BuildReport(templateDoc, xmlDataSource, "persons");

        // 6. Save the generated report.
        string reportPath = Path.Combine(workDir, "report.docx");
        templateDoc.Save(reportPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Report generated at: {reportPath}");
    }

    // Creates an XML file with a list of persons.
    private static void CreateSampleXml(string filePath)
    {
        // Use invariant culture when converting numbers to strings.
        var persons = new[]
        {
            new { Name = "Alice", Age = 30 },
            new { Name = "Bob",   Age = 45 },
            new { Name = "Carol", Age = 27 }
        };

        XElement root = new XElement("Persons",
            new XElement("Person",
                new XElement("Name", persons[0].Name),
                new XElement("Age", persons[0].Age.ToString(CultureInfo.InvariantCulture))),
            new XElement("Person",
                new XElement("Name", persons[1].Name),
                new XElement("Age", persons[1].Age.ToString(CultureInfo.InvariantCulture))),
            new XElement("Person",
                new XElement("Name", persons[2].Name),
                new XElement("Age", persons[2].Age.ToString(CultureInfo.InvariantCulture)))
        );

        // Save the XML document.
        XDocument doc = new XDocument(root);
        doc.Save(filePath);
    }

    // Creates a simple Word document with LINQ Reporting tags.
    private static void CreateTemplateDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a heading.
        builder.Writeln("People Report");
        builder.Writeln();

        // Insert a foreach loop over the collection "persons".
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age:  <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}
