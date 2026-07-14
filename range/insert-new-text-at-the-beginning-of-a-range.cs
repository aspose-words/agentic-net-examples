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
        builder.Writeln("Original text.");

        // Move the cursor to the very start of the document.
        builder.MoveToDocumentStart();

        // Insert new text at the beginning of the document's range.
        builder.Write("Inserted text ");

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
