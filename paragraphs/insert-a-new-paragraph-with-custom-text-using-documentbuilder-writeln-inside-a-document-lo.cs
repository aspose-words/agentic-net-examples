using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample data to be written as separate paragraphs.
        string[] paragraphs = new string[]
        {
            "This is the first inserted paragraph.",
            "Here is the second paragraph, added in a loop.",
            "Finally, the third paragraph completes the example."
        };

        // Loop through the array and write each string as a new paragraph.
        foreach (string text in paragraphs)
        {
            // Writeln inserts the text and then adds a paragraph break.
            builder.Writeln(text);
        }

        // Save the resulting document to the local file system.
        doc.Save("InsertedParagraphs.docx");
    }
}
