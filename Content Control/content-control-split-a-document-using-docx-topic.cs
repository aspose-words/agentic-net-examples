using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the source DOCX document that contains content controls (StructuredDocumentTag nodes).
        Document sourceDoc = new Document("SourceDocument.docx");

        // Get all content controls in the document.
        NodeCollection contentControls = sourceDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        // Iterate through each content control and save its contents as a separate DOCX file.
        for (int i = 0; i < contentControls.Count; i++)
        {
            // Cast the node to StructuredDocumentTag.
            StructuredDocumentTag sdt = (StructuredDocumentTag)contentControls[i];

            // Create a new empty document to hold the extracted part.
            Document partDoc = new Document();

            // Ensure the new document has a section (required for a valid DOCX).
            partDoc.RemoveAllChildren();
            partDoc.AppendChild(new Section(partDoc));

            // Import the content control node (including its children) into the new document.
            Node importedNode = partDoc.ImportNode(sdt, true);
            partDoc.FirstSection.Body.AppendChild(importedNode);

            // Build a filename based on the content control's title or index.
            string title = string.IsNullOrEmpty(sdt.Title) ? $"Part_{i + 1}" : sdt.Title.Replace(" ", "_");
            string outFileName = $"{title}.docx";

            // Save the extracted part as a separate DOCX file.
            partDoc.Save(outFileName);
        }
    }
}
