using System;
using Aspose.Words;
using Aspose.Words.Markup;

class XmlMappingRangedSdtExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting nodes.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a custom XML part that will be the data source for the content control.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = "<root><value>Mapped text from XML</value></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Insert the start of a ranged Structured Document Tag (content control).
        // Use RichText type so the control can contain paragraphs.
        StructuredDocumentTagRangeStart sdtStart = new StructuredDocumentTagRangeStart(doc, SdtType.RichText);
        builder.InsertNode(sdtStart);

        // Insert some placeholder content that will be inside the ranged content control.
        builder.Writeln("This paragraph is inside the ranged content control.");

        // Insert the end of the ranged Structured Document Tag.
        StructuredDocumentTagRangeEnd sdtEnd = new StructuredDocumentTagRangeEnd(doc, sdtStart.Id);
        builder.InsertNode(sdtEnd);

        // Map the content control to the XML element using XPath.
        // The control will display the text of the <value> element.
        sdtStart.XmlMapping.SetMapping(xmlPart, "/root[1]/value[1]", string.Empty);

        // Save the document to a file.
        doc.Save("RangedContentControlWithXmlMapping.docx");
    }
}
