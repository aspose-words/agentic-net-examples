using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the document that will be inserted.
        Document srcDoc = new Document("Source.docx");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Attach a DocumentBuilder to the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Move the cursor to the end of the document (where we want to insert).
        builder.MoveToDocumentEnd();

        // Optional: add a page break before the inserted content.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the source document, preserving its original formatting.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the result as a legacy .doc file using DocSaveOptions.
        DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
        dstDoc.Save("Combined.doc", saveOptions);
    }
}
