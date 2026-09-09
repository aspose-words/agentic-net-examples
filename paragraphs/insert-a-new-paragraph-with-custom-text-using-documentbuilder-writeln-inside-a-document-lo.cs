using System;
using Aspose.Words;

public class ParagraphInsertionExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder attached to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample data to be inserted as separate paragraphs.
        string[] paragraphs = new string[]
        {
            "First custom paragraph.",
            "Second custom paragraph.",
            "Third custom paragraph."
        };

        // Loop through the data and write each string as a new paragraph.
        foreach (string text in paragraphs)
        {
            // Writeln inserts the text and ends the paragraph.
            builder.Writeln(text);
        }

        // Save the document to the file system.
        doc.Save("InsertedParagraphs.docx");
    }
}
