using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertDocumentExample
{
    static void Main()
    {
        // Path to the source document that will be inserted.
        string sourcePath = "Source.docx";

        // Path where the resulting document will be saved.
        string resultPath = "Result.docx";

        // Load the source document (load rule).
        Document srcDoc = new Document(sourcePath);

        // Create a new blank destination document (create rule).
        Document dstDoc = new Document();

        // Initialize a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Insert an empty paragraph at the current cursor position (insert‑paragraph rule).
        // The builder's cursor is now positioned at the start of this new paragraph.
        builder.InsertParagraph();

        // Insert the previously loaded document at the cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted content.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document to disk (save rule).
        dstDoc.Save(resultPath);
    }
}
