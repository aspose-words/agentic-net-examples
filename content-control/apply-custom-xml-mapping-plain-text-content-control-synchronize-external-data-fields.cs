using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // 1. Create a new empty Word document.
        Document doc = new Document();

        // 2. Add a custom XML part that will act as the external data source.
        //    The XML contains two <text> elements – we will bind the content control to the second one.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = "<root><text>First value</text><text>Second value</text></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // 3. Create a plain‑text StructuredDocumentTag (content control) at block level.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);

        // 4. Map the content control to the desired XML node.
        //    XPath points to the second <text> element; no namespace prefixes are needed, so an empty string is passed.
        sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/text[2]", string.Empty);

        // 5. Insert the content control into the document body so it becomes visible in the file.
        doc.FirstSection.Body.AppendChild(sdt);

        // 6. Save the resulting document.
        doc.Save(@"C:\Artifacts\CustomXmlMapping.docx");
    }
}
