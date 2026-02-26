using System;
using Aspose.Words;

namespace AsposeWordsJoinExample
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document (destination).
            Document dstDoc = new Document();

            // 2. Load the source document that we want to insert.
            //    The file path should point to an existing DOCX file.
            Document srcDoc = new Document(@"C:\Docs\SourceDocument.docx");

            // 3. Use DocumentBuilder to position the cursor where the source will be inserted.
            //    Here we move to the end of the destination document's body.
            DocumentBuilder builder = new DocumentBuilder(dstDoc);
            builder.MoveToDocumentEnd();

            // 4. Insert the source document at the current cursor position.
            //    Keep the source formatting while inserting.
            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // 5. Save the combined document to a new file.
            dstDoc.Save(@"C:\Docs\CombinedDocument.docx");
        }
    }
}
