using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data with a "Country" attribute.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<persons>
    <person Name=""John Doe"" Country=""USA"" />
    <person Name=""Anna Smith"" Country=""Canada"" />
    <person Name=""Mike Johnson"" Country=""usa"" />
    <person Name=""Li Wei"" Country=""China"" />
</persons>";
        string xmlPath = "people.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // Create a LINQ Reporting template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a foreach tag that filters persons by Country using case‑insensitive comparison.
        // The LINQ Reporting engine does not support StringComparison, so we compare lower‑cased values.
        builder.Writeln("<<foreach [p in persons.Where(p => p.Country.ToLower() == \"usa\")]>>");
        builder.Writeln("Name: <<[p.Name]>>  Country: <<[p.Country]>>");
        builder.Writeln("<</foreach>>");

        // Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report. The data source name must match the root element name used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
