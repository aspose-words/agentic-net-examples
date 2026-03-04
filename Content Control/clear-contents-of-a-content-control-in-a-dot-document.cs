using System;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the DOT template document.
        Document doc = new Document("Template.dot");

        // Find every content control (StructuredDocumentTag) in the document.
        foreach (Node node in doc.GetChildNodes(NodeType.StructuredDocumentTag, true))
        {
            StructuredDocumentTag sdt = (StructuredDocumentTag)node;

            // Clear the contents of the content control.
            // If a placeholder is defined, it will be displayed after clearing.
            sdt.Clear();
        }

        // Save the updated document.
        doc.Save("Template_Cleared.dot");
    }
}
