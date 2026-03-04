using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.docx");

        // Retrieve all content controls (structured document tags) in the document.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        // Iterate through each content control.
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // If the control is a plain‑text content control, replace its inner text.
            if (sdt.SdtType == SdtType.PlainText)
            {
                // Remove any existing child nodes (the old text).
                sdt.RemoveAllChildren();

                // Use DocumentBuilder to insert new text at the position of the content control.
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.MoveTo(sdt);
                builder.Write("Updated content");
            }

            // Lock the contents so the user cannot edit the text.
            sdt.LockContents = true;

            // Lock the control itself so the user cannot delete it.
            sdt.LockContentControl = true;
        }

        // Ensure any fields in the document are refreshed.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
