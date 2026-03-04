using System;
using Aspose.Words;
using Aspose.Words.Markup;

class UpdateContentControls
{
    static void Main()
    {
        // Load the existing DOCX file.
        Document doc = new Document("Input.docx");

        // Retrieve all StructuredDocumentTag nodes (content controls) in the document.
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        // Iterate through each content control and replace its contents.
        foreach (StructuredDocumentTag sdt in sdtNodes)
        {
            // Remove any existing content inside the control.
            sdt.Clear();

            // Insert the new text that should appear inside the control.
            // You can customize the text based on the control's Title, Tag, etc.
            sdt.AppendChild(new Run(doc, "Updated value"));
        }

        // Save the modified document to a new file.
        doc.Save("Output.docx");
    }
}
