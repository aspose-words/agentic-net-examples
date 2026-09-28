using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add initial content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the original document text.");

        // Insert new text at the very beginning of the document.
        DocumentBuilder insertBuilder = new DocumentBuilder(doc);
        insertBuilder.MoveToDocumentStart();
        insertBuilder.Write("Inserted at start. ");

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
