using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

class UpdateContentControls
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Find all Structured Document Tags (content controls) in the document.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        // Iterate through each content control.
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // Move the builder's cursor to the start of the content control.
            builder.MoveTo(sdt);

            // Replace the existing content with new text.
            builder.Write("Updated content");

            // Example: lock the content control so the user cannot delete it.
            sdt.LockContentControl = true;

            // Example: lock the contents so the user cannot edit the text.
            sdt.LockContents = true;
        }

        // Update any fields that may be present in the document.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("Output.doc");
    }
}
