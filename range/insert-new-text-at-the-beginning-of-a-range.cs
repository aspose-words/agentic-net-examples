using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add initial content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph 1.");
        builder.Writeln("Original paragraph 2.");

        // Insert new text at the very beginning of the document's range.
        // Move the builder's cursor to the start of the document and write the text.
        builder.MoveToDocumentStart();
        builder.Write("Inserted at start. ");

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
