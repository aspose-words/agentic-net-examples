using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document with two paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");

        // Move the builder cursor to the start of the second paragraph.
        Paragraph secondParagraph = doc.FirstSection.Body.Paragraphs[1];
        builder.MoveTo(secondParagraph);

        // Insert an empty paragraph before the second paragraph.
        Paragraph emptyParagraph = builder.InsertParagraph();

        // Verify that the inserted paragraph is empty.
        if (emptyParagraph.GetText().Trim().Length != 0)
            throw new InvalidOperationException("The inserted paragraph is not empty.");

        // Save the resulting document.
        doc.Save("EmptyParagraphInserted.docx");
    }
}
