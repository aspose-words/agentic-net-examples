using Aspose.Words;
using Aspose.Words.Markup;

class Program
{
    static void Main()
    {
        // Load the DOT template.
        Document doc = new Document("Template.dot");

        // Retrieve the first content control (StructuredDocumentTag) in the document.
        StructuredDocumentTag sdt = (StructuredDocumentTag)doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

        // If a content control is found, clear its contents.
        if (sdt != null)
        {
            sdt.Range.Delete(); // Removes all characters inside the content control.
        }

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
