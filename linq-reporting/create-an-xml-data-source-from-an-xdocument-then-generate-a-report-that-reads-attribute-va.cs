using System;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create sample XML data that contains Order elements with attributes.
        // -----------------------------------------------------------------
        XDocument xml = new XDocument(
            new XElement("Orders",
                new XElement("Order",
                    new XAttribute("Id", "1"),
                    new XAttribute("Customer", "Alice")),
                new XElement("Order",
                    new XAttribute("Id", "2"),
                    new XAttribute("Customer", "Bob")),
                new XElement("Order",
                    new XAttribute("Id", "3"),
                    new XAttribute("Customer", "Charlie"))
            )
        );

        // Save the XML to a file that will be used as the data source.
        string xmlPath = Path.Combine(Environment.CurrentDirectory, "Orders.xml");
        xml.Save(xmlPath);

        // -----------------------------------------------------------------
        // 2. Build a Word template programmatically and insert LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document template = new Document();                     // blank document
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Order Report");                       // heading
        builder.Writeln();                                     // empty line

        // The root element name of the XML data source is "Orders".
        // Iterate over each Order element.
        builder.Writeln("<<foreach [order in Orders]>>");
        // Output the Id and Customer attribute values.
        builder.Writeln("Order Id: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.Customer]>>");
        builder.Writeln("<</foreach>>");

        // Save the template so it can be loaded for reporting.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and bind the XML data source.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // XmlDataSource reads the XML file; the root name ("Orders") will be used in the template.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report using the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "Orders");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        doc.Save(outputPath);
    }
}
