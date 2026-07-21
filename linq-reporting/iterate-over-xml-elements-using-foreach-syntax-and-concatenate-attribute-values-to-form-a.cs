using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable code pages for XML handling (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a blank Word document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Template: iterate over XML <item> elements and concatenate attribute values.
        builder.Writeln("<<foreach [item in items]>>");
        // Concatenate the 'name' and 'color' attributes with a hyphen.
        builder.Writeln("<<[item.name]>>-<<[item.color]>>");
        builder.Writeln("<</foreach>>");

        // Sample XML data with attributes.
        string xmlContent =
            @"<?xml version='1.0' encoding='utf-8'?>
              <items>
                  <item name='Apple'  color='Red' />
                  <item name='Banana' color='Yellow' />
                  <item name='Grape'  color='Purple' />
              </items>";

        // Load XML into a stream and create an XmlDataSource.
        using (MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent)))
        {
            XmlDataSource xmlDataSource = new XmlDataSource(xmlStream);

            // Build the report using the XML data source. The root name is "items".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, xmlDataSource, "items");
        }

        // Save the generated report.
        doc.Save("ReportOutput.docx");
    }
}
