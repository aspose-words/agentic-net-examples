using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public string Gender { get; set; } = "";
}

public class Model
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        const string xmlPath = "people.xml";
        File.WriteAllText(xmlPath,
            @"<People>
                <Person Name='Alice' Gender='Female' />
                <Person Name='Bob' Gender='Male' />
                <Person Name='Charlie' Gender='male' />
                <Person Name='Diana' Gender='FEMALE' />
              </People>");

        // Load XML and filter nodes where Gender attribute equals "male" (case‑insensitive).
        XDocument xdoc = XDocument.Load(xmlPath);
        List<Person> filtered = xdoc.Root!
            .Elements("Person")
            .Where(e => string.Equals((string?)e.Attribute("Gender"), "male", StringComparison.OrdinalIgnoreCase))
            .Select(e => new Person
            {
                Name = (string?)e.Attribute("Name") ?? "",
                Gender = (string?)e.Attribute("Gender") ?? ""
            })
            .ToList();

        // Wrap the filtered collection for the reporting engine.
        Model model = new Model { Persons = filtered };

        // Create a LINQ Reporting template programmatically.
        const string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Filtered persons (Gender = male):");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Gender]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template and build the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        const string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
