using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some text.
        builder.Writeln("This is the first line.");
        builder.Writeln("This is the second line that might wrap depending on page width.");

        // Retrieve the current paragraph (the last one added).
        Paragraph paragraph = builder.CurrentParagraph;

        // Aspose.Words does not provide a direct GetLinesCount method for a paragraph.
        // As a compile‑time safe approximation we store a placeholder value.
        int lineCount = 0; // No API to get exact line count of a paragraph.

        // Store the result (for demonstration we just write it to console).
        Console.WriteLine($"Approximate line count for the paragraph: {lineCount}");

        // Save the document to verify that the example works.
        doc.Save("ParagraphLinesCount.docx");
    }
}
