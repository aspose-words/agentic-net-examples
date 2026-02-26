using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();

        // Add a simple paragraph (optional, just to have content).
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello DOCM!");

        // Save the document in the macro‑enabled DOCM format.
        doc.Save("Result.docm", SaveFormat.Docm);
    }
}
