using System;
using Aspose.Words;
using Aspose.Words.Saving;

class JoinDocumentsExample
{
    static void Main()
    {
        // Load the main document (the one that will receive the inserted content).
        Document mainDoc = new Document("MainDocument.docx");

        // Load the document that we want to insert.
        Document insertDoc = new Document("DocumentToInsert.docx");

        // Create a DocumentBuilder for the main document.
        DocumentBuilder builder = new DocumentBuilder(mainDoc);

        // Move the cursor to the beginning of the desired section.
        // Sections are 0‑based; change the index to target a different section.
        int targetSectionIndex = 1; // example: insert after the first section
        builder.MoveToSection(targetSectionIndex);

        // Insert the entire content of insertDoc at the current cursor position.
        // KeepSourceFormatting preserves the original formatting of the inserted document.
        builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        mainDoc.Save("CombinedResult.docx");
    }
}
