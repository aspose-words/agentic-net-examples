using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model for the report.
    public class Person
    {
        public string Name { get; set; } = "";
        public string Role { get; set; } = "";
    }

    public static void Main()
    {
        // Prepare sample XML data.
        const string xmlFile = "people.xml";
        File.WriteAllText(xmlFile,
@"<people>
    <person name='John Doe' role='Admin' />
    <person name='Jane Smith' role='User' />
    <person name='Bob Johnson' role='admin' />
    <person name='Alice Brown' role='Guest' />
</people>");

        // Load XML and filter nodes where the 'role' attribute equals "admin" (case‑insensitive).
        XDocument xDoc = XDocument.Load(xmlFile);
        List<Person> filteredPersons = xDoc.Root!
            .Elements("person")
            .Where(p => string.Equals((string?)p.Attribute("role"), "admin", StringComparison.OrdinalIgnoreCase))
            .Select(p => new Person
            {
                Name = (string?)p.Attribute("name") ?? "",
                Role = (string?)p.Attribute("role") ?? ""
            })
            .ToList();

        // Create a template document with LINQ Reporting tags.
        const string templateFile = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Filtered persons (role = admin):");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Role: <<[p.Role]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templateFile);

        // Load the template and build the report using the filtered data.
        Document reportDoc = new Document(templateFile);
        ReportingEngine engine = new ReportingEngine();

        // The root object name must match the tag reference ("persons").
        engine.BuildReport(reportDoc, filteredPersons, "persons");

        // Save the final report.
        const string outputFile = "report.docx";
        reportDoc.Save(outputFile);
    }
}
