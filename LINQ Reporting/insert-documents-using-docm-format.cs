using System;
using Aspose.Words;
using Aspose.Words.Saving;

class InsertDocmExample
{
    static void Main()
    {
        // Path to the folder that contains input documents.
        // Replace with your actual directory.
        string inputFolder = @"C:\Docs\Input\";
        // Path to the folder where the result will be saved.
        // Replace with your actual directory.
        string outputFolder = @"C:\Docs\Output\";

        // Load the source document that will be inserted.
        Document srcDoc = new Document(inputFolder + "SourceDocument.docx");

        // Create a new blank destination document.
        Document dstDoc = new Document();

        // Use DocumentBuilder to position the cursor at the end of the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);
        builder.MoveToDocumentEnd();

        // Insert a page break before the inserted content (optional).
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the source document into the destination document.
        // KeepSourceFormatting preserves the original formatting of the source.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document as a macro‑enabled DOCM file.
        dstDoc.Save(outputFolder + "CombinedDocument.docm", SaveFormat.Docm);
    }
}
