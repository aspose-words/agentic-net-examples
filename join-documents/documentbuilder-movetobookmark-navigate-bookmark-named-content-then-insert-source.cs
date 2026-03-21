using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace InsertDocAtBookmarkExample
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create (or load) the destination document that contains the bookmark.
            // -----------------------------------------------------------------
            Document dstDoc = new Document();                     // creates a blank document
            DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);

            // For demonstration purposes we create a bookmark named "Content".
            // In a real scenario the bookmark could already exist in a loaded document.
            dstBuilder.StartBookmark("Content");
            dstBuilder.Write("Placeholder text that will be replaced by the inserted document.");
            dstBuilder.EndBookmark("Content");

            // -----------------------------------------------------------------
            // 2. Create the source document that we want to insert.
            // -----------------------------------------------------------------
            Document srcDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
            srcBuilder.Writeln("This is the first paragraph of the source document.");
            srcBuilder.Writeln("This is the second paragraph of the source document.");

            // -----------------------------------------------------------------
            // 3. Move the builder's cursor to the bookmark named "Content".
            // -----------------------------------------------------------------
            bool bookmarkFound = dstBuilder.MoveToBookmark("Content");
            if (!bookmarkFound)
                throw new InvalidOperationException("Bookmark 'Content' was not found in the destination document.");

            // -----------------------------------------------------------------
            // 4. Insert the source document at the bookmark position,
            //    preserving the source formatting.
            // -----------------------------------------------------------------
            dstBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // -----------------------------------------------------------------
            // 5. Save the resulting document.
            // -----------------------------------------------------------------
            dstDoc.Save("Result.docx", SaveFormat.Docx);
        }
    }
}
