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

        // Define XML content and a unique identifier for the custom XML part.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = "<root><author>John Doe</author><title>Sample Document</title></root>";

        // Add the custom XML part to the document's collection.
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Create a plain‑text content control (structured document tag) for the author.
        StructuredDocumentTag authorTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
        // Bind the content control to the <author> element in the custom XML part.
        authorTag.XmlMapping.SetMapping(xmlPart, "/root[1]/author[1]", string.Empty);
        authorTag.Title = "Author";

        // Insert the author content control into the document body.
        doc.FirstSection.Body.AppendChild(authorTag);

        // Create another content control for the title.
        StructuredDocumentTag titleTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);
        titleTag.XmlMapping.SetMapping(xmlPart, "/root[1]/title[1]", string.Empty);
        titleTag.Title = "Title";

        // Insert the title content control into the document body.
        doc.FirstSection.Body.AppendChild(titleTag);

        // Save the document as a DOCX file.
        string outputPath = "ContentControlBinding.docx";
        doc.Save(outputPath);
    }
}
