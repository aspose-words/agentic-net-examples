using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add two initial paragraphs.
        builder.Writeln("Paragraph 1: This is the first paragraph.");
        builder.Writeln("Paragraph 2: This is the second paragraph.");

        // Locate the second paragraph node.
        Paragraph secondParagraph = doc.FirstSection.Body.Paragraphs[1];

        // Move the builder's cursor to the start of the second paragraph.
        builder.MoveTo(secondParagraph);

        // Insert an empty paragraph at this exact position.
        // The InsertParagraph method inserts a paragraph break and returns the new (empty) paragraph.
        Paragraph emptyParagraph = builder.InsertParagraph();

        // Optionally, add more content after the inserted empty paragraph.
        builder.Writeln("Paragraph 3: This follows the empty paragraph.");

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
