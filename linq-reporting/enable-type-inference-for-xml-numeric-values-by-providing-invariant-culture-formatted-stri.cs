using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data with numeric values formatted using invariant culture (dot as decimal separator).
        const string xmlFileName = "data.xml";
        string xmlContent =
            @"<?xml version=""1.0"" encoding=""utf-8""?>
<orders>
    <order>
        <Id>1</Id>
        <Amount>1234.56</Amount>
        <Quantity>10</Quantity>
    </order>
    <order>
        <Id>2</Id>
        <Amount>7890.12</Amount>
        <Quantity>5</Quantity>
    </order>
</orders>";
        File.WriteAllText(xmlFileName, xmlContent);

        // Create a Word template programmatically.
        const string templateFileName = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        // Iterate directly over the collection represented by the root element "orders".
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Amount: <<[order.Amount]>>");
        builder.Writeln("Quantity: <<[order.Quantity]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templateFileName);

        // Load the template for reporting.
        var doc = new Document(templateFileName);

        // Load XML data source.
        var xmlDataSource = new XmlDataSource(xmlFileName);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "orders");

        // Save the generated report.
        const string outputFileName = "report.docx";
        doc.Save(outputFileName);
    }
}
