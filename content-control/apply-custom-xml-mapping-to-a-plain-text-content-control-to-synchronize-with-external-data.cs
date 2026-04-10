using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading paragraph.
        builder.Writeln("Custom XML Mapping Example");

        // Create a custom XML part that holds external data.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = @"<root>
    <field1>First value</field1>
    <field2>Second value</field2>
</root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Insert a plain‑text content control that will display <field1>.
        StructuredDocumentTag sdtField1 = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block)
        {
            Title = "Field1Control",
            Tag = "field1"
        };
        sdtField1.XmlMapping.SetMapping(xmlPart, "/root[1]/field1[1]", string.Empty);
        // Append the content control to the document body (valid block container).
        doc.FirstSection.Body.AppendChild(sdtField1);
        // Add an empty paragraph after the control to act as a line break.
        doc.FirstSection.Body.AppendChild(new Paragraph(doc));

        // Insert another plain‑text content control that will display <field2>.
        StructuredDocumentTag sdtField2 = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block)
        {
            Title = "Field2Control",
            Tag = "field2"
        };
        sdtField2.XmlMapping.SetMapping(xmlPart, "/root[1]/field2[1]", string.Empty);
        doc.FirstSection.Body.AppendChild(sdtField2);
        doc.FirstSection.Body.AppendChild(new Paragraph(doc));

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "CustomXmlMapping.docx");
        doc.Save(outputPath);
    }
}
