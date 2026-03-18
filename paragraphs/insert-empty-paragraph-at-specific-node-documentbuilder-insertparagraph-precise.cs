using System;
using Aspose.Words;

namespace InsertEmptyParagraphExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Initialize a DocumentBuilder attached to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some initial content so we have a reference node.
            builder.Writeln("First paragraph.");
            builder.Writeln("Second paragraph.");

            // Locate the paragraph after which we want to insert an empty paragraph.
            // In this example we insert after the first paragraph.
            Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];

            // Move the builder's cursor to the end of the target paragraph.
            // When the cursor is positioned on a Paragraph, InsertParagraph inserts a new
            // paragraph break right after it (i.e., between this paragraph and the next one).
            builder.MoveTo(firstParagraph);

            // Insert an empty paragraph at the current cursor position.
            // The method returns the newly created Paragraph, which is also the builder's CurrentParagraph.
            Paragraph emptyParagraph = builder.InsertParagraph();

            // Optionally, you can verify that the inserted paragraph is indeed empty.
            // It will contain no runs and its text will be just the paragraph break.
            Console.WriteLine("Inserted paragraph is empty: " + (emptyParagraph.Runs.Count == 0));

            // Save the document to a file.
            doc.Save("InsertedEmptyParagraph.docx");
        }
    }
}
