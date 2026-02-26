using System;
using Aspose.Words;
using Aspose.Words.Markup; // Added for SdtType enum
using Aspose.Words.Saving;

class InsertDocumentIntoContentControl
{
    static void Main()
    {
        // Paths – adjust as needed.
        string sourcePath = @"C:\Docs\Source.docx";
        string resultPath = @"C:\Docs\InsertDocIntoContentControl.docx";

        // Create a new blank document.
        Document mainDoc = new Document();

        // Create a DocumentBuilder attached to the new document.
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        // Insert a rich‑text content control (StructuredDocumentTag) where the other document will be placed.
        builder.InsertStructuredDocumentTag(SdtType.RichText);
        // Move the cursor inside the newly created content control.
        builder.MoveToStructuredDocumentTag(0, 0);

        // Load the document that we want to insert.
        Document srcDoc = new Document(sourcePath);

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the original styles of the inserted document.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Optionally add some text after the inserted content.
        builder.Writeln();
        builder.Writeln("Insertion completed.");

        // Save the resulting document.
        mainDoc.Save(resultPath);
    }
}
