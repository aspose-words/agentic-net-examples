using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertDocxIntoHeaderBookmark
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the destination document with a header that contains a bookmark.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(destDoc);

        // Ensure the document has at least one section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.StartBookmark("HeaderInsert");
        builder.Write("Placeholder text that will be replaced.");
        builder.EndBookmark("HeaderInsert");

        // Return the builder to the main story (optional, not required for insertion).
        builder.MoveToDocumentEnd();

        // -----------------------------------------------------------------
        // 2. Create the source document that will be inserted.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the content from the source document.");
        srcBuilder.Writeln("It will be inserted at the bookmark inside the header.");

        // -----------------------------------------------------------------
        // 3. Move the cursor to the bookmark inside the header.
        // -----------------------------------------------------------------
        bool found = builder.MoveToBookmark("HeaderInsert");
        if (!found)
            throw new InvalidOperationException("Bookmark 'HeaderInsert' not found in the header.");

        // -----------------------------------------------------------------
        // 4. Insert the source document at the bookmark position.
        // -----------------------------------------------------------------
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Save the modified document.
        // -----------------------------------------------------------------
        string outPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        destDoc.Save(outPath, SaveFormat.Docx);

        Console.WriteLine($"Document saved to: {outPath}");
    }
}
