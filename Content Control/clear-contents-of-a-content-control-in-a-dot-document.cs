using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the DOT template document.
        Document doc = new Document("Template.dot");

        // Locate the first content control (StructuredDocumentTag) in the document.
        StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);
        if (sdt != null)
        {
            // Clear the contents of the content control.
            // If a placeholder is defined, it will be displayed after clearing.
            sdt.Clear();
        }

        // Save the updated document.
        doc.Save("Result.dot");
    }
}
