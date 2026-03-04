using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the destination document (the one that will receive the insertion).
        Document destination = new Document("Destination.docx");

        // Create a DocumentBuilder for the destination document.
        DocumentBuilder builder = new DocumentBuilder(destination);

        // Create an empty Run that will act as the insertion point.
        Run insertionRun = new Run(destination, string.Empty);

        // Insert the Run into the document at the current cursor position.
        builder.InsertNode(insertionRun);

        // Move the builder's cursor to the newly inserted Run.
        builder.MoveTo(insertionRun);

        // Load the source document that we want to insert.
        Document source = new Document("Source.docx");

        // Insert the source document at the Run position.
        // KeepSourceFormatting preserves the original formatting of the source.
        builder.InsertDocument(source, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        destination.Save("Result.docx");
    }
}
