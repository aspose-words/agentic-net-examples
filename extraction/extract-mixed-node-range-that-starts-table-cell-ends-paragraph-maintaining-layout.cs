using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ExtractMixedRange
{
    static void Main()
    {
        // Load the source document.
        Document srcDoc = new Document("Input.docx");

        // Locate the starting Cell (first cell in the first table for this example).
        Cell startCell = (Cell)srcDoc.GetChild(NodeType.Cell, 0, true);

        // Locate the ending Paragraph (first paragraph after the table for this example).
        // Adjust the index as needed to point to the desired paragraph.
        Paragraph endParagraph = (Paragraph)srcDoc.GetChild(NodeType.Paragraph, 5, true);

        // Create a new empty document that will hold the extracted range.
        Document destDoc = new Document();
        // Ensure the document has a section and a body to receive nodes.
        destDoc.RemoveAllChildren();
        Section section = new Section(destDoc);
        destDoc.AppendChild(section);
        Body body = new Body(destDoc);
        section.AppendChild(body);

        // Iterate from the start cell to the end paragraph, importing each node into the new document.
        Node current = startCell;
        while (current != null)
        {
            // Import the node (deep clone) into the destination document.
            Node importedNode = destDoc.ImportNode(current, true);
            body.AppendChild(importedNode);

            // Stop when we have appended the ending paragraph.
            if (current == endParagraph)
                break;

            // Move to the next node in the document order.
            current = current.NextSibling;
        }

        // Save the extracted range as a separate document, preserving layout.
        destDoc.Save("ExtractedRange.docx");
    }
}
