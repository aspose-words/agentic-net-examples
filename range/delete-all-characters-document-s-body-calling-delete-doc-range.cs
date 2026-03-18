using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new document (or load an existing one)
        Document doc = new Document(); // For an existing file use: new Document("input.docx");

        // Delete all characters in the document's range (which includes the body)
        doc.Range.Delete();

        // Save the modified document
        doc.Save("output.docx");
    }
}
