using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the destination (template) document.
        Document dstDoc = new Document("Template.docx");

        // Load the document that will be inserted at a bookmark.
        Document srcInsert = new Document("Insert.docx");

        // Load the document that will be appended to the end.
        Document srcAppend = new Document("Append.docx");

        // Insert the srcInsert document at the bookmark named "InsertHere".
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.MoveToBookmark("InsertHere");
        builder.InsertDocument(srcInsert, ImportFormatMode.KeepSourceFormatting);

        // Append the srcAppend document to the end of the destination document.
        dstDoc.AppendDocument(srcAppend, ImportFormatMode.UseDestinationStyles);

        // Save the combined document.
        dstDoc.Save("Result.docx");
    }
}
