using System;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML data with attributes
        XDocument xdoc = new XDocument(
            new XElement("Orders",
                new XElement("Order",
                    new XAttribute("Id", "1"),
                    new XAttribute("Customer", "Alice"),
                    new XAttribute("Amount", "100.5")),
                new XElement("Order",
                    new XAttribute("Id", "2"),
                    new XAttribute("Customer", "Bob"),
                    new XAttribute("Amount", "200.0"))
            )
        );

        // Save XML to a file (optional, just for demonstration)
        string xmlPath = "data.xml";
        xdoc.Save(xmlPath);

        // Create a Word template programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Orders Report");
        builder.Writeln("==============");
        builder.Writeln();
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order Id: <<[order.Id]>>, Customer: <<[order.Customer]>>, Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template (simulating a separate load step)
        Document doc = new Document(templatePath);

        // Load XML data source
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "Orders");

        // Save the generated report
        string outputPath = "report.docx";
        doc.Save(outputPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
