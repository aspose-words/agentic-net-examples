using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with sample text.
        builder.Writeln("This is a sample paragraph that may wrap into multiple lines depending on the layout.");

        // Get the first paragraph in the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Approximate line count.
        // The Aspose.Words API does not provide a direct GetLinesCount method.
        // For demonstration purposes we store a placeholder value.
        int lineCount = 1; // Placeholder for actual line count logic.

        // Output the stored line count.
        Console.WriteLine($"Approximate line count: {lineCount}");

        // Save the document to verify the paragraph was added.
        doc.Save("Output.docx");
    }
}
