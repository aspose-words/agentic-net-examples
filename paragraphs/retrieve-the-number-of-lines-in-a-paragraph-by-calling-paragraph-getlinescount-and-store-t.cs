using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a single paragraph that contains explicit line breaks.
        // The "\r" characters represent line breaks within the paragraph.
        builder.Writeln("First line.\rSecond line.\rThird line.");

        // Retrieve the first (and only) paragraph in the document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Approximate the number of lines by counting line‑break characters in the paragraph text.
        // GetText() returns the paragraph text including the terminating paragraph mark.
        string text = paragraph.GetText();

        // Remove the trailing paragraph mark before splitting.
        text = text.TrimEnd('\r', '\n');

        // Split on any line‑break sequence to count lines.
        int lineCount = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None).Length;

        // Store the result (here we simply output it).
        Console.WriteLine($"Approximate line count: {lineCount}");

        // Save the document to demonstrate the full workflow.
        doc.Save("Output.docx");
    }
}
