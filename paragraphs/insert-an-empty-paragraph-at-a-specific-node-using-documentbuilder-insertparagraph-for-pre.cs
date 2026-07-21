using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add two initial paragraphs.
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");

        // Move the builder's cursor to the first paragraph.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        builder.MoveTo(firstParagraph);

        // Insert an empty paragraph right after the first paragraph.
        // The returned Paragraph object represents the newly inserted empty paragraph.
        Paragraph emptyParagraph = builder.InsertParagraph();

        // Save the resulting document.
        doc.Save("InsertedParagraph.docx");
    }
}
