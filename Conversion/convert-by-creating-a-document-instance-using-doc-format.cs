using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();

        // Save the document in the legacy DOC format (Microsoft Word 97‑2007).
        doc.Save("Result.doc", SaveFormat.Doc);
    }
}
