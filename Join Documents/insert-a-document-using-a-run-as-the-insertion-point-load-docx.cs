using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the destination document (the one that will receive the insertion).
        Document dstDoc = new Document("Destination.docx");

        // Load the source document (the document to be inserted).
        Document srcDoc = new Document("Source.docx");

        // Locate a Run node in the destination document to serve as the insertion point.
        // Here we simply take the first Run found in the document.
        Run insertionRun = (Run)dstDoc.GetChild(NodeType.Run, 0, true);

        // Create a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Move the builder's cursor to the selected Run.
        builder.MoveTo(insertionRun);

        // Insert the source document at the current cursor position.
        // KeepSourceFormatting preserves the formatting of the inserted content.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the modified document.
        dstDoc.Save("Result.docx");
    }
}
