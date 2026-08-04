using System;
using System.IO;
using Aspose.Words;

namespace ParagraphInsertionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add initial content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("First paragraph.");   // Paragraph 0
            builder.Writeln("Second paragraph."); // Paragraph 1

            // Move the builder's cursor to the end of the first paragraph.
            // This positions the cursor right before the second paragraph.
            builder.MoveTo(doc.FirstSection.Body.Paragraphs[0]);

            // Insert an empty paragraph at the current cursor position.
            // The method returns the newly inserted Paragraph, which will be empty.
            Paragraph emptyParagraph = builder.InsertParagraph();

            // (Optional) Verify that the inserted paragraph has no child nodes (i.e., it is empty).
            // This line does not affect the document; it is just for demonstration.
            bool isEmpty = emptyParagraph.HasChildNodes == false;

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "InsertedEmptyParagraph.docx");
            doc.Save(outputPath);
        }
    }
}
