using System;
using Aspose.Words;
using Aspose.Words.Markup;

class DuplicateContentControl
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Source.docx");

        // Find the first content control (StructuredDocumentTag) in the document.
        StructuredDocumentTag originalTag = (StructuredDocumentTag)doc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)[0];

        // Deep clone the content control together with its child nodes.
        StructuredDocumentTag clonedTag = (StructuredDocumentTag)originalTag.Clone(true);

        // Prepare a DocumentBuilder to insert the cloned control.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the location where the copy should be inserted.
        // Use a bookmark named "InsertHere" if it exists; otherwise fall back to the document end.
        if (doc.Range.Bookmarks["InsertHere"] != null)
        {
            builder.MoveToBookmark("InsertHere");
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        // Insert the cloned content control at the current cursor position.
        builder.InsertNode(clonedTag);

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
