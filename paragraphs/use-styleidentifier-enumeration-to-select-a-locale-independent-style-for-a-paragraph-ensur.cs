using System;
using Aspose.Words;

namespace ParagraphStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Insert a paragraph with some text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This paragraph uses the Normal style.");

            // Apply a locale‑independent built‑in style using StyleIdentifier.
            Paragraph paragraph = doc.LastSection.Body.LastParagraph;
            paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

            // Save the document.
            doc.Save("StyledParagraph.docx");
        }
    }
}
