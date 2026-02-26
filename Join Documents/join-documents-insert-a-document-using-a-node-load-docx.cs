using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a destination document (blank document)
        // -------------------------------------------------
        Document dstDoc = new Document();                     // lifecycle: create
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // -------------------------------------------------
        // 2. Define a bookmark that marks the insertion point
        // -------------------------------------------------
        builder.StartBookmark("InsertHere");
        builder.Write("Text before insertion. ");
        builder.EndBookmark("InsertHere");
        builder.Write(" Text after insertion.");

        // -------------------------------------------------
        // 3. Load the source document that will be inserted
        // -------------------------------------------------
        Document srcDoc = new Document("Source.docx");        // lifecycle: load

        // -------------------------------------------------
        // 4. Move the builder cursor to the bookmark and insert the source document
        // -------------------------------------------------
        builder.MoveToBookmark("InsertHere");
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting); // feature: InsertDocument

        // -------------------------------------------------
        // 5. Save the combined document
        // -------------------------------------------------
        dstDoc.Save("Result.docx");                          // lifecycle: save
    }
}
