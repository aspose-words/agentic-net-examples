using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample texts to insert as separate paragraphs.
        string[] texts = { "First paragraph.", "Second paragraph.", "Third paragraph." };

        // Insert each text as a new paragraph using Writeln inside a loop.
        foreach (string text in texts)
        {
            builder.Writeln(text);
        }

        // Save the document.
        doc.Save("Output.docx");
    }
}
