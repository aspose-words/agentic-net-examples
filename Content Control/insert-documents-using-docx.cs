using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsInsertDocumentExample
{
    class Program
    {
        static void Main()
        {
            // Create a destination document (blank document).
            Document dstDoc = new Document();

            // Use DocumentBuilder to add some initial content.
            DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
            dstBuilder.Writeln("This is the destination document.");

            // Insert a page break before the insertion point (optional).
            dstBuilder.InsertBreak(BreakType.PageBreak);

            // Create a source document that will be inserted.
            Document srcDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
            srcBuilder.Writeln("This is the source document being inserted.");
            srcBuilder.Writeln("It can contain multiple paragraphs, tables, images, etc.");

            // Move the cursor to the end of the destination document.
            dstBuilder.MoveToDocumentEnd();

            // Insert the source document at the current cursor position.
            // KeepSourceFormatting preserves the original formatting of the source.
            dstBuilder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the combined document to a DOCX file.
            dstDoc.Save("CombinedDocument.docx", SaveFormat.Docx);
        }
    }
}
