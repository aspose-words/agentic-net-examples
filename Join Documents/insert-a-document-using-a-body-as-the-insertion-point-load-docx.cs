using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a blank destination document.
        Document dstDoc = new Document();

        // Initialize a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(dstDoc);

        // Write a paragraph that will serve as the insertion point.
        builder.Writeln("=== Insertion Point ===");

        // Keep a reference to the paragraph just written.
        Paragraph insertionParagraph = builder.CurrentParagraph;

        // Load the source DOCX document to be inserted.
        Document srcDoc = new Document("Source.docx");

        // Move the builder's cursor to the insertion paragraph.
        builder.MoveTo(insertionParagraph);

        // Insert the source document at the cursor position, preserving its formatting.
        builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // Add a paragraph after the inserted content.
        builder.Writeln("=== Insertion Completed ===");

        // Save the combined document.
        dstDoc.Save("Result.docx");
    }
}
