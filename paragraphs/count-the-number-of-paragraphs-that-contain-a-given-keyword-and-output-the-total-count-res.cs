using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Define the keyword to search for.
        const string keyword = "Aspose";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs.
        builder.Writeln("This is the first paragraph.");
        builder.Writeln($"This paragraph mentions {keyword} in the text.");
        builder.Writeln("Another paragraph without the keyword.");
        builder.Writeln($"{keyword} appears again here.");
        builder.Writeln("Final paragraph.");

        // Save the document (optional, demonstrates lifecycle compliance).
        doc.Save("SampleDocument.docx");

        // Count paragraphs that contain the keyword.
        int count = 0;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // Get the paragraph text (includes the paragraph mark).
            string text = para.GetText();

            // Check if the text contains the keyword (case‑insensitive).
            if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                count++;
            }
        }

        // Output the result.
        Console.WriteLine($"Number of paragraphs containing \"{keyword}\": {count}");
    }
}
