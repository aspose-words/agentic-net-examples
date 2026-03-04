using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the documents that will be inserted/appended.
        Document srcInsert = new Document("SourceInsert.docx");
        Document srcAppend = new Document("SourceAppend.docx");

        // Create a new blank destination document.
        Document dst = new Document();

        // Prepare a bookmark in the destination where the first document will be inserted.
        DocumentBuilder builder = new DocumentBuilder(dst);
        builder.StartBookmark("InsertPoint");
        builder.Writeln("Text before insertion.");
        builder.EndBookmark("InsertPoint");
        builder.Writeln("Text after insertion.");

        // Move the cursor to the bookmark and insert the first document.
        builder.MoveToBookmark("InsertPoint");
        builder.InsertDocument(srcInsert, ImportFormatMode.KeepSourceFormatting);

        // Move the cursor to the end of the document and append the second document.
        builder.MoveToDocumentEnd();
        dst.AppendDocument(srcAppend, ImportFormatMode.UseDestinationStyles);

        // Save the combined document.
        dst.Save("Combined.docx");
    }
}
