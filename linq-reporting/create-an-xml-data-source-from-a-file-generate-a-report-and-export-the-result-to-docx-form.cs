using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a sample XML data file.
        const string xmlFileName = "data.xml";
        string xmlContent =
@"<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <Total>123.45</Total>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <Total>67.89</Total>
    </Order>
</Orders>";
        File.WriteAllText(xmlFileName, xmlContent);

        // Create the LINQ Reporting template programmatically.
        const string templateFileName = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        // Iterate over the collection of Order elements.
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Total: $<<[order.Total]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templateFileName);

        // Load the template for report generation.
        Document doc = new Document(templateFileName);

        // Load XML data source.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlFileName);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "Orders");

        // Save the generated report.
        const string outputFileName = "report.docx";
        doc.Save(outputFileName);
    }
}
