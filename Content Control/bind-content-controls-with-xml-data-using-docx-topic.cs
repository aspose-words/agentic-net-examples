using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // XML data that will be bound to content controls.
        string xml = @"<root>
    <title>Example Title</title>
    <author>John Doe</author>
    <content>This is the body text.</content>
</root>";

        // Add the XML as a custom XML part to the document.
        string partId = Guid.NewGuid().ToString("B");
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(partId, xml);

        // Title content control.
        StructuredDocumentTag titleTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
        titleTag.Title = "Title";
        titleTag.XmlMapping.SetMapping(xmlPart, "/root[1]/title[1]", string.Empty);
        doc.FirstSection.Body.AppendChild(titleTag);

        // Author content control.
        StructuredDocumentTag authorTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
        authorTag.Title = "Author";
        authorTag.XmlMapping.SetMapping(xmlPart, "/root[1]/author[1]", string.Empty);
        doc.FirstSection.Body.AppendChild(authorTag);

        // Content body control.
        StructuredDocumentTag bodyTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
        bodyTag.Title = "Content";
        bodyTag.XmlMapping.SetMapping(xmlPart, "/root[1]/content[1]", string.Empty);
        doc.FirstSection.Body.AppendChild(bodyTag);

        // Save the resulting DOCX file.
        doc.Save("ContentControlsMapped.docx");
    }
}
