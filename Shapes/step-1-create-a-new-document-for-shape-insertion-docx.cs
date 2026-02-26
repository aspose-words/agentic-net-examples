using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has the minimum required nodes (section, body, paragraph).
        doc.EnsureMinimum();

        // Save the document in DOCX format.
        doc.Save("ShapeInsertion.docx");
    }
}
