using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertDocxIntoHeaderBookmark
{
    static void Main()
    {
        // Path to the destination document that contains a header with a bookmark.
        string destPath = @"C:\Docs\Destination.docx";

        // Path to the source document that will be inserted.
        string srcPath = @"C:\Docs\Source.docx";

        // Load the destination document.
        Document destDoc = new Document(destPath);

        // Load the source document that we want to insert.
        Document srcDoc = new Document(srcPath);

        // Create a DocumentBuilder attached to the destination document.
        DocumentBuilder builder = new DocumentBuilder(destDoc);

        // -----------------------------------------------------------------
        // 1. Move the builder's cursor to the primary header of the first
        //    section. All subsequent operations will affect the header.
        // -----------------------------------------------------------------
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // -----------------------------------------------------------------
        // 2. (Optional) Ensure a bookmark exists in the header.
        //    If the header already contains a bookmark, this step can be
        //    omitted. Here we create a bookmark named "HeaderInsert".
        // -----------------------------------------------------------------
        builder.StartBookmark("HeaderInsert");
        builder.Write("Placeholder text that will be replaced.");
        builder.EndBookmark("HeaderInsert");

        // -----------------------------------------------------------------
        // 3. Move the cursor to the start of the bookmark where the source
        //    document will be inserted.
        // -----------------------------------------------------------------
        bool found = builder.MoveToBookmark("HeaderInsert");
        if (!found)
            throw new InvalidOperationException("Bookmark 'HeaderInsert' not found in the header.");

        // -----------------------------------------------------------------
        // 4. Insert the source document at the bookmark position.
        //    KeepSourceFormatting preserves the original formatting of the
        //    inserted document, including any header/footer styles it may
        //    contain.
        // -----------------------------------------------------------------
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // 5. Save the modified document as DOCX, preserving header formatting.
        //    The default OoxmlSaveOptions keep the existing compliance level.
        // -----------------------------------------------------------------
        string outPath = @"C:\Docs\Result.docx";
        destDoc.Save(outPath, SaveFormat.Docx);
    }
}
