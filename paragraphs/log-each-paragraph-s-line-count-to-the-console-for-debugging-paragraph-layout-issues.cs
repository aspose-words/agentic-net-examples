using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First paragraph – single line.");
        builder.Writeln("Second paragraph – line one.\nLine two.");
        builder.Writeln("Third paragraph – line one.\nLine two.\nLine three.");

        // Save the document (optional, just to have an output file).
        doc.Save("SampleOutput.docx");

        // Retrieve all paragraph nodes in the document.
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true)
                           .OfType<Paragraph>()
                           .ToList();

        // Log an approximate line count for each paragraph.
        for (int i = 0; i < paragraphs.Count; i++)
        {
            // Approximate line count by counting newline characters in the paragraph text.
            // This is a safe compile‑time approximation because Aspose.Words does not expose a direct line‑count API.
            string text = paragraphs[i].GetText(); // Includes the paragraph break character at the end.
            int lineCount = text.Count(c => c == '\n') + 1; // Add 1 for the last line (or the only line).

            Console.WriteLine($"Paragraph {i + 1}: Approximate line count = {lineCount}");
        }
    }
}
