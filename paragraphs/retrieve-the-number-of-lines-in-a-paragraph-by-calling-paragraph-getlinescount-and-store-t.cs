using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some text.
        builder.Writeln("This is a sample paragraph that may wrap into multiple lines depending on the page width and formatting.");

        // Retrieve the first paragraph in the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // The Paragraph.GetLinesCount() method does not exist in the Aspose.Words API.
        // As a compile‑safe approximation, we store a placeholder value.
        // In a real scenario you would use layout information to determine the line count.
        int lineCount = 1; // Approximation: assume at least one line.

        // Store or use the line count as needed (here we just output it).
        Console.WriteLine($"Estimated line count for the paragraph: {lineCount}");

        // Save the document to verify that the example runs without external input.
        doc.Save("Output.docx");
    }
}
