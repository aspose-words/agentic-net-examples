using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new (empty) document.
        Document doc = new Document();

        // Add a custom XML part that will hold the data we want to bind to.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlContent = "<root><title>Sample Title</title><value>123</value></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Create a plain‑text content control (StructuredDocumentTag).
        StructuredDocumentTag contentControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);

        // Bind the content control to the <title> element of the custom XML part.
        // The XPath points to the first <title> element under the first <root> element.
        bool isMapped = contentControl.XmlMapping.SetMapping(xmlPart, "/root[1]/title[1]", string.Empty);
        if (!isMapped)
            throw new InvalidOperationException("Failed to map the content control to the XML part.");

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(contentControl);

        // Save the document as a macro‑enabled DOCM file.
        doc.Save("Output.docm", SaveFormat.Docm);
    }
}
