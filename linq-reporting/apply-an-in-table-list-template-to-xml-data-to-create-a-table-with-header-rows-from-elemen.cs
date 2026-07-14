using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data
        string xmlPath = "data.xml";
        XDocument xmlDoc = new XDocument(
            new XElement("Items",
                new XElement("Item",
                    new XElement("Name", "John Doe"),
                    new XElement("Age", "30"),
                    new XElement("Country", "USA")),
                new XElement("Item",
                    new XElement("Name", "Jane Smith"),
                    new XElement("Age", "25"),
                    new XElement("Country", "Canada"))
            )
        );
        xmlDoc.Save(xmlPath);

        // Extract column names from the first Item element
        List<string> columnNames = xmlDoc.Root?
            .Elements("Item")
            .FirstOrDefault()?
            .Elements()
            .Select(e => e.Name.LocalName)
            .ToList() ?? new List<string>();

        // Create the template document programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin foreach over Items
        builder.Writeln("<<foreach [item in Items]>>");

        // Start table
        Table table = builder.StartTable();

        // Header row
        foreach (string col in columnNames)
        {
            builder.InsertCell();
            builder.Writeln(col);
        }
        builder.EndRow();

        // Data row (inside foreach)
        foreach (string col in columnNames)
        {
            builder.InsertCell();
            builder.Writeln($"<<[item.{col}]>>");
        }
        builder.EndRow();

        // End table and foreach
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template
        string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Build the report using XmlDataSource
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, new XmlDataSource(xmlPath), "Items");

        // Save the generated report
        string outputPath = Path.Combine("output", "report.docx");
        Directory.CreateDirectory("output");
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
