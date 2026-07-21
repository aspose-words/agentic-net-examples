using System;
using Aspose.Words;

public class ParagraphLineCountExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some sample text.
        builder.Writeln("This is the first line of the paragraph.");
        builder.Writeln("This is the second line of the same paragraph.");
        builder.Writeln("And this is the third line.");

        // The Paragraph class does not expose a GetLinesCount method.
        // As a compile‑safe approximation we can update the document layout
        // and retrieve the total number of lines in the document.
        // For a single‑paragraph document this value approximates the paragraph's line count.
        doc.UpdateWordCount(true);
        int approximateParagraphLineCount = doc.BuiltInDocumentProperties.Lines;

        // Store the result (for demonstration purposes we simply output it).
        Console.WriteLine("Approximate line count for the paragraph: " + approximateParagraphLineCount);

        // Save the document to verify that the example runs without errors.
        doc.Save("ParagraphLineCountExample.docx");
    }
}
