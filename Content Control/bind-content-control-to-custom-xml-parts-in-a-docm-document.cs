using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class ContentControlWithCustomXml
{
    static void Main()
    {
        // Folder where the output document will be saved.
        string outputFolder = @"C:\Temp\";
        Directory.CreateDirectory(outputFolder);

        // Create a new blank document.
        Document doc = new Document();

        // Define a unique identifier for the custom XML part.
        string xmlPartId = Guid.NewGuid().ToString("B");

        // XML data that will be stored in the custom XML part.
        string xmlContent = "<root><title>Sample Title</title><body>Sample body text.</body></root>";

        // Add the custom XML part to the document's collection.
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlContent);

        // Create a plain‑text content control (structured document tag).
        StructuredDocumentTag contentControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);

        // Bind the content control to the <title> element of the custom XML part.
        // The XPath points to the first <title> element under the first <root> element.
        bool isMapped = contentControl.XmlMapping.SetMapping(xmlPart, "/root[1]/title[1]", string.Empty);

        // Optional: verify that the mapping succeeded.
        if (!isMapped)
            throw new InvalidOperationException("Failed to map the content control to the custom XML part.");

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(contentControl);

        // Save the document as a macro‑enabled DOCM file.
        string outputPath = Path.Combine(outputFolder, "ContentControlWithCustomXml.docm");
        doc.Save(outputPath, SaveFormat.Docm);
    }
}
