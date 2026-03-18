using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Get all StructuredDocumentTag nodes in the document (including nested ones).
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

        // Iterate over a copy of the collection to avoid modification issues during removal.
        foreach (StructuredDocumentTag sdt in sdtNodes.Cast<StructuredDocumentTag>().ToList())
        {
            // If the tag contains any child nodes, remove the tag itself but keep its content.
            // RemoveSelfOnly moves the children to the parent, effectively placing the tag body
            // outside the markup element.
            if (sdt.GetChildNodes(NodeType.Any, false).Count > 0)
                sdt.RemoveSelfOnly();
        }

        // Save the processed document.
        doc.Save("Output.docx");
    }
}
