using System;
using Aspose.Words;

namespace AsposeWordsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for inserting content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Apply a locale‑independent style using StyleIdentifier.
            // Heading1 is a built‑in style that works regardless of the document language.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("This paragraph uses the built‑in Heading1 style.");

            // Switch to another built‑in style (Normal) using the same approach.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This paragraph uses the Normal style.");

            // Save the document to the current working directory.
            string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "StyledParagraphs.docx");
            doc.Save(outputPath);
        }
    }
}
