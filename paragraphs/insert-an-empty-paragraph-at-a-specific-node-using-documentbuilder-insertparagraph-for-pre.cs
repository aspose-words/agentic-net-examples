using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial paragraphs.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Locate the first paragraph in the document.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;

        // Move the builder to the first paragraph.
        builder.MoveTo(firstParagraph);

        // Insert an empty paragraph immediately after the first paragraph.
        Paragraph emptyParagraph = builder.InsertParagraph();

        // Verify that the inserted paragraph is indeed empty.
        // (Trim is used to ignore any hidden end-of-paragraph marker.)
        if (emptyParagraph.GetText().Trim().Length == 0)
        {
            // The paragraph is empty as expected.
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
