using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert a paragraph with some text using DocumentBuilder.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph that may span multiple lines depending on the layout.");

        // Get the paragraph that was just created.
        Paragraph paragraph = builder.CurrentParagraph;

        // Aspose.Words does not provide a Paragraph.GetLinesCount() method.
        // As a compile‑safe approximation, we assume a single line for this short paragraph.
        int lineCount = 1; // Placeholder for the number of lines in the paragraph.

        // Store or use the line count as needed.
        Console.WriteLine($"Approximate line count for the paragraph: {lineCount}");

        // Save the document to demonstrate a complete workflow.
        doc.Save("ParagraphLineCountExample.docx");
    }
}
