using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data with attributes.
        const string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<items>
    <item Name=""Apple"" Price=""1.20"" />
    <item Name=""Banana"" Price=""0.80"" />
    <item Name=""Cherry"" Price=""2.50"" />
</items>";
        const string xmlPath = "data.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // Create a Word template that uses LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Iterate over the XML elements and concatenate attribute values.
        // The data source name will be "items", so we iterate directly over "items".
        builder.Writeln("<<foreach [item in items]>>");
        builder.Writeln("<<[item.Name]>> - <<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report using the data source named "items".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, dataSource, "items");

        // Save the generated report.
        reportDoc.Save("output.docx");
    }
}
