using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample texts to be inserted as separate paragraphs.
        string[] texts = { "First custom paragraph.", "Second custom paragraph.", "Third custom paragraph." };

        // Insert each text inside the loop using Writeln, which adds the text and a paragraph break.
        foreach (string text in texts)
        {
            builder.Writeln(text);
        }

        // Save the resulting document.
        doc.Save("InsertedParagraphs.docx");
    }
}
