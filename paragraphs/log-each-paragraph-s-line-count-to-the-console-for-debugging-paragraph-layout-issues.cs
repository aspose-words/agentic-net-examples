using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a few sample paragraphs.
        builder.Writeln("First short paragraph.");
        builder.Writeln("Second paragraph contains a bit more text to demonstrate how the approximation works. " +
                        "It will be split into multiple lines based on the character count per line.");
        builder.Writeln("Third paragraph is even longer. " +
                        "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut " +
                        "labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco " +
                        "laboris nisi ut aliquip ex ea commodo consequat.");

        // Approximate line count for each paragraph.
        // Since Aspose.Words does not expose a per‑paragraph line count, we estimate it by assuming a fixed
        // number of characters per line (e.g., 80). This provides a compile‑safe placeholder for debugging.
        const int charsPerLine = 80;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        for (int i = 0; i < paragraphs.Count; i++)
        {
            Paragraph para = (Paragraph)paragraphs[i];
            string text = para.GetText(); // Includes the paragraph break character.
            int approxLines = Math.Max(1, (text.Length + charsPerLine - 1) / charsPerLine);
            Console.WriteLine($"Paragraph {i + 1}: Approximate line count = {approxLines}");
        }

        // Save the document (optional, demonstrates the full lifecycle).
        doc.Save("ParagraphLineCounts.docx");
    }
}
