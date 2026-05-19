using System;
using Aspose.Words;

namespace AsposeWordsStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder for easy content insertion.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set a locale‑independent style using StyleIdentifier.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("This paragraph uses Heading 1 style.");

            // Change to another built‑in style.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("This paragraph uses Normal style.");

            // Save the document to the current directory.
            string outputPath = "StyledParagraph.docx";
            doc.Save(outputPath);
        }
    }
}
