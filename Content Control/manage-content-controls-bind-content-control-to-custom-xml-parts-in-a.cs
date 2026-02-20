using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

class ContentControlBindingExample
{
    static void Main()
    {
        // Create a new blank document (DOCM will be set on save).
        Document doc = new Document();

        // Define a custom XML part with a unique identifier and some sample XML content.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = "<root><title>Sample Title</title><value>12345</value></root>";

        // Add the custom XML part to the document's collection.
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Create a plain‑text content control (structured document tag) at block level.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);

        // Map the content control to the <value> element of the custom XML part using XPath.
        // The third parameter is for namespace prefix mappings; empty string means none.
        sdt.XmlMapping.SetMapping(xmlPart, "/root[1]/value[1]", string.Empty);

        // Insert the content control into the first section's body.
        doc.FirstSection.Body.AppendChild(sdt);

        // Save the document as a macro‑enabled DOCM file.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docm);
        doc.Save("ContentControlBinding.docm", saveOptions);
    }
}
