using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Markup;

class ApplySectionBackgroundWithContentControls
{
    static void Main()
    {
        // Load an existing document.
        Document doc = new Document("Input.docx");

        // Iterate over all sections in the document.
        foreach (Section section in doc.Sections)
        {
            // Determine a background color based on the section index (example expression).
            int sectionIndex = doc.Sections.IndexOf(section);
            Color background = GetBackgroundColorForSection(sectionIndex);

            // Create a block‑level rich‑text content control (SDT) that will wrap the whole section.
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                // Set the color of the content control frame (this is the property that can be set).
                Color = background,
                // Optional: make the control visible as tags.
                Appearance = SdtAppearance.Tags
            };

            // Preserve the original nodes of the section.
            List<Node> originalNodes = new List<Node>();
            foreach (Node node in section.Body.GetChildNodes(NodeType.Any, true))
                originalNodes.Add(node);

            // Remove the original nodes from the body and add them to the content control.
            foreach (Node node in originalNodes)
            {
                section.Body.RemoveChild(node);
                sdt.AppendChild(node);
            }

            // Insert the content control as the sole child of the section body.
            section.Body.PrependChild(sdt);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }

    // Example evaluated expression: alternating colors based on section index.
    private static Color GetBackgroundColorForSection(int index)
    {
        // Even sections get LightGray, odd sections get LightBlue.
        return (index % 2 == 0) ? Color.LightGray : Color.LightBlue;
    }
}
