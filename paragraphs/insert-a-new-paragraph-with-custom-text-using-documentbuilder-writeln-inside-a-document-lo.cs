using System;
using Aspose.Words;

namespace ParagraphInsertionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Sample texts to insert as separate paragraphs.
            string[] customTexts = { "First custom paragraph.", "Second custom paragraph.", "Third custom paragraph." };

            // Loop through the texts and insert each as a new paragraph using Writeln.
            foreach (string text in customTexts)
            {
                // Writeln writes the text and starts a new paragraph.
                builder.Writeln(text);
            }

            // Save the resulting document.
            doc.Save("InsertedParagraphs.docx");
        }
    }
}
