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
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            // For demonstration purposes we create a bookmark named "Content".
            // In a real scenario the bookmark could already exist in a loaded document.
            builder.StartBookmark("Content");
            builder.Write("Placeholder text that will be replaced by the inserted document.");
            builder.EndBookmark("Content");

            // -----------------------------------------------------------------
            // 2. Load the source document that we want to insert.
            // -----------------------------------------------------------------
            Document srcDoc = new Document("Source.docx");        // load existing DOCX file

            // -----------------------------------------------------------------
            // 3. Move the builder's cursor to the bookmark named "Content".
            // -----------------------------------------------------------------
            bool bookmarkFound = builder.MoveToBookmark("Content");
            if (!bookmarkFound)
                throw new InvalidOperationException("Bookmark 'Content' was not found in the destination document.");

            // -----------------------------------------------------------------
            // 4. Insert the source document at the bookmark position,
            //    preserving the source formatting.
            // -----------------------------------------------------------------
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // -----------------------------------------------------------------
            // 5. Save the resulting document.
            // -----------------------------------------------------------------
            dstDoc.Save("Result.docx", SaveFormat.Docx);
        }
    }
}
