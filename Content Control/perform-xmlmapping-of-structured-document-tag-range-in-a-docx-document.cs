using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // -----------------------------------------------------------------
        // 1. Add a custom XML part that will hold the data to be mapped.
        // -----------------------------------------------------------------
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlPartContent = "<root><text>First element</text><text>Second element</text></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlPartContent);

        // -----------------------------------------------------------------
        // 2. Insert a ranged Structured Document Tag (SDT) start and end.
        // -----------------------------------------------------------------
        // The start node represents the beginning of the range.
        StructuredDocumentTagRangeStart rangeStart = new StructuredDocumentTagRangeStart(doc, SdtType.PlainText);
        // The end node must have the same Id as the start node.
        StructuredDocumentTagRangeEnd rangeEnd = new StructuredDocumentTagRangeEnd(doc, rangeStart.Id);

        // Place the range at the beginning of the first section's body.
        // Insert the start node before the first paragraph.
        doc.FirstSection.Body.InsertBefore(rangeStart, doc.FirstSection.Body.FirstParagraph);
        // Insert the end node immediately after the start node.
        doc.FirstSection.Body.InsertAfter(rangeEnd, rangeStart);

        // -----------------------------------------------------------------
        // 3. Map the range start to a specific XML node in the custom part.
        // -----------------------------------------------------------------
        // The XPath points to the second <text> element inside <root>.
        rangeStart.XmlMapping.SetMapping(xmlPart, "/root[1]/text[2]", null);

        // -----------------------------------------------------------------
        // 4. Save the document.
        // -----------------------------------------------------------------
        doc.Save("StructuredDocumentTagRangeStartXmlMapping.docx");
    }
}
