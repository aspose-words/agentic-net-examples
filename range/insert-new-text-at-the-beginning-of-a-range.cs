using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a paragraph with initial text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original text.");

        // Move the builder to the very start of the document (the beginning of the document's range).
        builder.MoveToDocumentStart();

        // Insert the new text at that position.
        builder.Write("Inserted text. ");

        // Save the document.
        doc.Save("Result.docx");
    }
}
