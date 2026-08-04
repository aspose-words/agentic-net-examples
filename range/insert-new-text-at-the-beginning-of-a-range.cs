using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add initial content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("World!");

        // Insert new text at the very beginning of the document's range.
        // Move the builder's cursor to the start of the document and write the text.
        builder.MoveToDocumentStart();
        builder.Write("Hello ");

        // Optional verification (can be removed in production).
        string result = doc.GetText().Trim();
        if (result != "Hello World!")
            throw new InvalidOperationException("Text insertion failed.");

        // Save the modified document.
        doc.Save("InsertedText.docx");
    }
}
