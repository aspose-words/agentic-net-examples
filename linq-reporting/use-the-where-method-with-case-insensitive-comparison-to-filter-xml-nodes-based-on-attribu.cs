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
    public string City { get; set; } = "";
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample XML data with a "city" attribute.
        string xmlContent = @"
<people>
    <person name='Alice' city='London' />
    <person name='Bob' city='Paris' />
    <person name='Charlie' city='london' />
    <person name='Diana' city='New York' />
</people>";

        // Load XML into XDocument.
        XDocument xdoc = XDocument.Parse(xmlContent);

        // Filter persons where the city attribute equals "London" (case‑insensitive).
        var filtered = xdoc.Root!
            .Elements("person")
            .Where(p => string.Equals((string?)p.Attribute("city"), "London", StringComparison.OrdinalIgnoreCase))
            .Select(p => new Person
            {
                Name = (string?)p.Attribute("name") ?? "",
                City = (string?)p.Attribute("city") ?? ""
            })
            .ToList();

        // Prepare the model for the reporting engine.
        ReportModel model = new ReportModel { Persons = filtered };

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("People filtered by city (case‑insensitive):");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.City]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template and build the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}
