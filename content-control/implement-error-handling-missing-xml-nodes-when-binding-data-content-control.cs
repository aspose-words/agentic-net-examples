using System;
using Aspose.Words;
using Aspose.Words.Markup;

class ContentControlXmlBinding
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a custom XML part containing a simple element.
        string xmlPartId = Guid.NewGuid().ToString("B");
        string xml = "<root><item>Value1</item></root>";
        CustomXmlPart xmlPart = doc.CustomXmlParts.Add(xmlPartId, xml);

        // Create a plain‑text content control (structured document tag).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block);

        // -----------------------------------------------------------------
        // 1. Map to an existing XML node – this should succeed.
        // -----------------------------------------------------------------
        string existingXPath = "/root[1]/item[1]";
        bool mapped = sdt.XmlMapping.SetMapping(xmlPart, existingXPath, string.Empty);
        if (!mapped || !sdt.XmlMapping.IsMapped)
        {
            // Unexpected failure – provide a placeholder.
            SetPlaceholder(sdt, $"Node not found: {existingXPath}");
        }

        // -----------------------------------------------------------------
        // 2. Attempt to map to a missing XML node – this will fail.
        // -----------------------------------------------------------------
        string missingXPath = "/root[1]/nonexistent[1]";
        bool mappedMissing = sdt.XmlMapping.SetMapping(xmlPart, missingXPath, string.Empty);
        if (!mappedMissing || !sdt.XmlMapping.IsMapped)
        {
            // Handle the missing node gracefully by showing a placeholder message.
            SetPlaceholder(sdt, $"Missing XML node: {missingXPath}");
        }

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(sdt);

        // Save the resulting document.
        doc.Save("ContentControlWithErrorHandling.docx");
    }

    // Helper method that clears the control's contents and inserts a placeholder run.
    private static void SetPlaceholder(StructuredDocumentTag sdt, string message)
    {
        // Remove any existing child nodes.
        sdt.RemoveAllChildren();

        // Add a Run node containing the placeholder text.
        sdt.AppendChild(new Run(sdt.Document, message));

        // Ensure the placeholder is displayed when the control is empty.
        sdt.IsShowingPlaceholderText = true;
    }
}
