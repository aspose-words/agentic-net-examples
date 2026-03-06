using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new paragraph containing the desired text.
        builder.Writeln("This is a new paragraph inserted via Aspose.Words.");

        // Save the document to the file system.
        doc.Save("InsertedParagraph.docx");
    }
}
