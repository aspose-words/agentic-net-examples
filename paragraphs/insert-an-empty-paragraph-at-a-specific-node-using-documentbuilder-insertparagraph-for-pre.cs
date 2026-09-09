using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add two paragraphs with sample text.
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");

        // Move the builder cursor to the first paragraph (index 0).
        builder.MoveTo(doc.FirstSection.Body.Paragraphs[0]);

        // Insert an empty paragraph at the current cursor position.
        Paragraph emptyParagraph = builder.InsertParagraph();

        // Verify that the inserted paragraph is empty (its text is just a paragraph break).
        if (emptyParagraph.GetText() == "\r")
        {
            // The paragraph is correctly empty.
        }

        // Save the resulting document.
        doc.Save("InsertedEmptyParagraph.docx");
    }
}
