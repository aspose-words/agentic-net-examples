using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Path to the source document that already contains a ranged structured document tag.
        string inputPath = @"C:\Docs\Multi-section structured document tags.docx";

        // Path where the resulting document will be saved.
        string outputPath = @"C:\Output\StructuredDocumentTagRangeStartXmlMapping.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Create a custom XML part that will be used for mapping.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xmlPartContent = "<root><text>Text element #1</text><text>Text element #2</text></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xmlPartContent);

        // Verify that the XML part was added correctly (optional).
        if (Encoding.UTF8.GetString(xmlPart.Data) != xmlPartContent)
            throw new InvalidOperationException("Custom XML part content mismatch.");

        // Retrieve the first ranged structured document tag start node in the document.
        StructuredDocumentTagRangeStart sdtRangeStart = (StructuredDocumentTagRangeStart)doc.GetChild(
            NodeType.StructuredDocumentTagRangeStart, 0, true);

        // Map the structured document tag range start to the second <text> element in the custom XML part.
        // The XPath points to the desired node; the third parameter (namespace mappings) is optional.
        sdtRangeStart.XmlMapping.SetMapping(xmlPart, "/root[1]/text[2]", null);

        // Save the modified document.
        doc.Save(outputPath);
    }
}
