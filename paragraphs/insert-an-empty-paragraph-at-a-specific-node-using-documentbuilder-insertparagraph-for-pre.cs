using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add initial paragraphs.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Locate the first paragraph node where we want to insert the empty paragraph.
        Paragraph referenceParagraph = doc.FirstSection.Body.FirstParagraph;

        // Move the builder's cursor to the reference paragraph.
        builder.MoveTo(referenceParagraph);

        // Insert an empty paragraph after the reference paragraph.
        // InsertParagraph inserts a new paragraph after the current node and returns it.
        Paragraph emptyParagraph = builder.InsertParagraph();

        // Verify that the inserted paragraph is empty (its text should be just a paragraph break).
        string insertedText = emptyParagraph.GetText(); // Expected to be "\r"
        Console.WriteLine("Inserted paragraph is empty: " + (insertedText == "\r"));

        // Save the document to verify the result.
        doc.Save("Output.docx");
    }
}
