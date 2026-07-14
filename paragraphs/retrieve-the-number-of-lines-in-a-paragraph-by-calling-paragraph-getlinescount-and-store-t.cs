using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with enough text to wrap onto multiple lines.
        builder.Writeln("This is a long paragraph that is expected to span multiple lines when rendered in the document layout. " +
                        "It contains enough words to cause line wrapping in the final document.");

        // Get the paragraph we have just created.
        Paragraph paragraph = builder.CurrentParagraph;

        // Aspose.Words does not provide a Paragraph.GetLinesCount() method.
        // As a compile‑safe approximation we will store a placeholder value.
        int lineCount = 0;

        // Update the document layout so that built‑in line statistics are calculated.
        doc.UpdateWordCount(true);

        // Retrieve the total number of lines in the document (not per paragraph).
        // This demonstrates storing a line count value, even though it is not the exact paragraph line count.
        int totalDocumentLines = doc.BuiltInDocumentProperties.Lines;

        // Store the (approximate) line count.
        lineCount = totalDocumentLines;

        // Output the stored line count.
        Console.WriteLine($"Stored line count (approximation): {lineCount}");

        // Save the document to the file system.
        doc.Save("ParagraphLines.docx");
    }
}
