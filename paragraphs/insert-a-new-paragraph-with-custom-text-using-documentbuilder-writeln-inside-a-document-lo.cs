using System;
using Aspose.Words;

namespace ParagraphInsertionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();

            // Use DocumentBuilder to add some initial paragraphs.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Original Paragraph 1");
            builder.Writeln("Original Paragraph 2");
            builder.Writeln("Original Paragraph 3");

            // Loop through the existing paragraphs and insert a new paragraph after each one.
            // Note: The collection changes as we insert, so we increment the index manually.
            for (int i = 0; i < doc.FirstSection.Body.Paragraphs.Count; i++)
            {
                Paragraph currentParagraph = doc.FirstSection.Body.Paragraphs[i];

                // Move the builder to the end of the current paragraph.
                builder.MoveTo(currentParagraph);
                
                // Insert a new paragraph with custom text.
                builder.Writeln($"Inserted after paragraph {i + 1}");

                // Skip the newly inserted paragraph to avoid processing it again.
                i++;
            }

            // Save the resulting document.
            doc.Save("Result.docx");
        }
    }
}
