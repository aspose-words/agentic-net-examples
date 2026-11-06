using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial paragraphs.
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Locate the first paragraph node.
        Paragraph firstParagraph = doc.FirstSection.Body.Paragraphs[0];

        // Position the builder at the first paragraph.
        builder.MoveTo(firstParagraph);

        // Insert an empty paragraph right after the first paragraph.
        Paragraph emptyParagraph = builder.InsertParagraph();

        // Save the resulting document.
        doc.Save("Output.docx");

        // Indicate completion.
        Console.WriteLine("Empty paragraph inserted after the first paragraph. Document saved as Output.docx");
    }
}
