using System;
using Aspose.Words;

public class ParagraphKeywordCounter
{
    public static void Main()
    {
        // Keyword to search for
        string keyword = "Aspose";

        // Create a new document and a builder to add content
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample paragraphs
        builder.Writeln("This is a sample paragraph without the keyword.");
        builder.Writeln("Aspose.Words is a powerful library for document processing.");
        builder.Writeln("Another line that mentions Aspose.");
        builder.Writeln("No keyword here.");
        builder.Writeln("The keyword Aspose appears again.");

        // Save the document (optional output)
        doc.Save("ParagraphKeywordCount.docx");

        // Count paragraphs that contain the keyword (case‑insensitive)
        int count = 0;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            string text = para.GetText();
            if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                count++;
        }

        // Output the total count
        Console.WriteLine($"Number of paragraphs containing \"{keyword}\": {count}");
    }
}
