using System;
using Aspose.Words;

namespace AsposeWordsInsertAtBookmark
{
    class Program
    {
        static void Main()
        {
            // Load the destination document that contains a bookmark named "InsertHere".
            Document destDoc = new Document("Destination.docx");

            // Load the source document whose entire content will be inserted.
            Document srcDoc = new Document("Source.docx");

            // Create a DocumentBuilder for the destination document.
            DocumentBuilder builder = new DocumentBuilder(destDoc);

            // Move the builder's cursor to the bookmark.
            // Returns true if the bookmark exists; otherwise false.
            if (!builder.MoveToBookmark("InsertHere"))
                throw new InvalidOperationException("Bookmark 'InsertHere' not found in destination document.");

            // Insert the source document at the bookmark position.
            // KeepSourceFormatting preserves the original formatting of the source document.
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document.
            destDoc.Save("Combined.docx");
        }
    }
}
