using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample XML using XDocument.
        XDocument xDoc = new XDocument(
            new XElement("people",
                new XElement("person",
                    new XAttribute("name", "Alice"),
                    new XAttribute("age", "30")),
                new XElement("person",
                    new XAttribute("name", "Bob"),
                    new XAttribute("age", "25"))
            )
        );

        // Write the XML to a memory stream – this will be used as the XmlDataSource.
        using MemoryStream xmlStream = new();
        xDoc.Save(xmlStream);
        xmlStream.Position = 0; // Reset before reading.

        // Create the reporting template programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // The XML root (<people>) is treated as a collection of <person> elements,
        // so we iterate directly over the data source itself.
        builder.Writeln("<<foreach [p in data]>>");
        // Attribute values must be referenced with the attribute name inside double quotes.
        builder.Writeln("Name: <<[p.@\"name\"]>>");
        builder.Writeln("Age: <<[p.@\"age\"]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the XML data source.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlStream);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(template, xmlDataSource, "data");

        // Save the generated report.
        template.Save("Report.docx");
    }
}
