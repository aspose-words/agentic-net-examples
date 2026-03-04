using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the destination document that already contains a bookmark.
        Document destDoc = new Document("Destination.docx");

        // Create a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destDoc);

        // Move the builder's cursor to the bookmark named "InsertHere".
        // The method returns true if the bookmark exists.
        bool bookmarkFound = builder.MoveToBookmark("InsertHere");
        if (!bookmarkFound)
        {
            Console.WriteLine("Bookmark 'InsertHere' not found in the destination document.");
            return;
        }

        // Load the source document whose contents will be inserted.
        Document srcDoc = new Document("Source.docx");

        // Insert the source document at the current cursor position (the bookmark).
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the resulting document.
        destDoc.Save("Result.docx");
    }
}
