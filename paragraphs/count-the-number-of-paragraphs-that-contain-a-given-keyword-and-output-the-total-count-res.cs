using System;
using Aspose.Words;

public class ParagraphKeywordCounter
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs.
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("Aspose.Words is a powerful library for document processing.");
        builder.Writeln("This paragraph contains the keyword: Aspose.");
        builder.Writeln("Another line without the key term.");
        builder.Writeln("Keyword appears again: Aspose.");

        // Define the keyword to search for.
        string keyword = "Aspose";

        // Count paragraphs that contain the keyword (case‑insensitive).
        int count = 0;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // GetText includes the paragraph break; Trim removes it.
            string text = para.GetText().Trim();
            if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                count++;
        }

        // Output the total count.
        Console.WriteLine($"Paragraphs containing \"{keyword}\": {count}");
    }
}
