using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom XML part that will supply data for the content control.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xml = "<root><author>John Doe</author><title>Sample Book</title></root>";
        CustomXmlPart customXml = doc.CustomXmlParts.Add(xmlPartId, xml);

        // Use DocumentBuilder to insert nodes into the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the start of a ranged content control (plain‑text type).
        StructuredDocumentTagRangeStart rangeStart = new StructuredDocumentTagRangeStart(doc, SdtType.PlainText);
        rangeStart.Title = "BookInfo";               // optional title
        builder.InsertNode(rangeStart);               // place the start tag at the cursor

        // Content that will be bound to the XML mapping.
        builder.Writeln("Author: ");
        builder.Writeln("Title: ");

        // Insert the matching end tag for the ranged content control.
        StructuredDocumentTagRangeEnd rangeEnd = new StructuredDocumentTagRangeEnd(doc, rangeStart.Id);
        builder.InsertNode(rangeEnd);

        // Map the start tag to the <author> element of the custom XML part.
        // The XPath selects the first <author> element under <root>.
        rangeStart.XmlMapping.SetMapping(customXml, "/root[1]/author[1]", string.Empty);

        // Save the resulting document.
        doc.Save("MappedRangeContentControl.docx");
    }
}
