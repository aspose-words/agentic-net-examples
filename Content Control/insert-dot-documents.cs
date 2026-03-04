using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOT template file as a Document.
        Document dotDoc = new Document("Template.dot");

        // Create a new blank document that will receive the inserted content.
        Document resultDoc = new Document();

        // Attach a DocumentBuilder to the result document.
        DocumentBuilder builder = new DocumentBuilder(resultDoc);

        // Position the builder at the end of the document (or any desired location).
        builder.MoveToDocumentEnd();

        // Optional: insert a page break before the inserted content.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the DOT document into the result document, preserving its original formatting.
        builder.InsertDocument(dotDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the final document.
        resultDoc.Save("Combined.docx");
    }
}
