using System;
using Aspose.Words;

namespace ParagraphNavigationExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder for the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a few sample paragraphs.
            builder.Writeln("Paragraph 0: Introduction.");
            builder.Writeln("Paragraph 1: Overview.");
            builder.Writeln("Paragraph 2: Details.");
            builder.Writeln("Paragraph 3: Conclusion.");

            // Move the builder's cursor to the third paragraph (index 2, zero‑based).
            // characterIndex = 0 positions the cursor at the start of the paragraph.
            builder.MoveToParagraph(2, 0);

            // Apply formatting to the paragraph we have moved to.
            // Here we center‑align the paragraph text.
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Optionally, insert a new line after the formatted paragraph.
            builder.Writeln("This line was added after moving to paragraph index 2.");

            // Save the document to the current directory.
            string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "Result.docx");
            doc.Save(outputPath);
        }
    }
}
