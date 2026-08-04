using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for .NET Core environments.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample XML data file.
        string xmlPath = "orders.xml";
        File.WriteAllText(xmlPath, GetSampleXml());

        // Create the LINQ Reporting template and save it.
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "Orders");

        // Save the generated report.
        doc.Save("report.docx");
    }

    // Generates a simple template that repeats a single‑row block for each Order element.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin the foreach loop over the Orders collection.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Amount: <<[order.Amount]>>");
        // Insert a page break so each order starts on a new page (separate section).
        builder.InsertBreak(BreakType.PageBreak);
        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    // Returns a small XML document containing a list of orders.
    private static string GetSampleXml()
    {
        return @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
    <Order>
        <Id>1001</Id>
        <CustomerName>John Doe</CustomerName>
        <Amount>250.00</Amount>
    </Order>
    <Order>
        <Id>1002</Id>
        <CustomerName>Jane Smith</CustomerName>
        <Amount>175.50</Amount>
    </Order>
    <Order>
        <Id>1003</Id>
        <CustomerName>Bob Johnson</CustomerName>
        <Amount>320.75</Amount>
    </Order>
</Orders>";
    }
}
