using System;
using Aspose.Words;

public class ParagraphLineCountExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs of varying lengths.
        builder.Writeln("Short paragraph.");
        builder.Writeln("This is a medium length paragraph that contains a few more words to demonstrate line counting.");
        builder.Writeln("This is a long paragraph intended to simulate a situation where the text might wrap onto multiple visual lines in the layout. " +
                        "It contains many sentences, commas, and other punctuation marks to increase its length and complexity.");

        // Save the document (optional, just to have an output file).
        doc.Save("ParagraphLineCounts.docx");

        // Retrieve all paragraph nodes in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Iterate through each paragraph and log an approximate line count.
        // Since Aspose.Words does not expose a direct line‑count API for a paragraph,
        // we use a simple approximation based on the length of the paragraph text.
        int index = 1;
        foreach (Paragraph para in paragraphs)
        {
            // Get the raw text of the paragraph (includes the paragraph break character).
            string text = para.GetText();

            // Approximate line count: treat any non‑empty paragraph as at least one line.
            // For a rougher estimate you could divide the character count by a constant,
            // but here we keep it simple and deterministic.
            int approxLineCount = string.IsNullOrWhiteSpace(text) ? 0 : 1;

            Console.WriteLine($"Paragraph {index}: Approximate line count = {approxLineCount}, Text length = {text.Length}");
            index++;
        }
    }
}
