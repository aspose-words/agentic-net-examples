using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs with varying line breaks.
        builder.Writeln("First paragraph line one.");
        builder.Writeln("Second paragraph line one.\r\nSecond line.\r\nThird line.");
        builder.Writeln("Third paragraph with no line breaks.");

        // Save the document (optional for debugging).
        doc.Save("ParagraphLineCounts.docx");

        // Traverse all paragraphs and log an approximate line count.
        int paragraphIndex = 0;
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Approximate line count by counting non‑empty lines in the paragraph text.
            string text = paragraph.GetText(); // Includes paragraph mark.
            int lineCount = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Length;

            Console.WriteLine($"Paragraph {paragraphIndex}: {lineCount} line(s)");
            paragraphIndex++;
        }
    }
}
