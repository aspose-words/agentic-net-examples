using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Public property that matches the template tag.
    public string Name { get; set; } = string.Empty;
}

public class DataModel
{
    // Collection of persons to iterate over in the template.
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create sample XML with a default namespace.
        // -----------------------------------------------------------------
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<root xmlns=""urn:test"">
    <person>
        <name>Alice</name>
    </person>
    <person>
        <name>Bob</name>
    </person>
</root>";
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "people.xml");
        File.WriteAllText(xmlPath, xmlContent);

        // -----------------------------------------------------------------
        // 2. Load the XML into a strongly‑typed data model.
        // -----------------------------------------------------------------
        XDocument doc = XDocument.Load(xmlPath);
        XNamespace ns = "urn:test";

        DataModel model = new DataModel();

        foreach (XElement personElem in doc.Root!.Elements(ns + "person"))
        {
            string name = personElem.Element(ns + "name")?.Value ?? string.Empty;
            model.Persons.Add(new Person { Name = name });
        }

        // -----------------------------------------------------------------
        // 3. Build the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags that reference the model.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("<<[p.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template and generate the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // The root object name ("model") must match the name used in the template.
        engine.BuildReport(report, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(reportPath);
    }
}
