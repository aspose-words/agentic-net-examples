using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs to the document.
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("Aspose.Words is a powerful library.");
        builder.Writeln("Another line without the keyword.");
        builder.Writeln("Learning Aspose can be fun.");
        builder.Writeln("Final paragraph.");

        // Define the keyword to search for.
        string keyword = "Aspose";

        // Count the paragraphs that contain the keyword (case‑insensitive).
        int count = 0;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // Get the full text of the paragraph (includes the paragraph break).
            string text = para.GetText();

            if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                count++;
        }

        // Output the total count.
        Console.WriteLine($"Paragraphs containing \"{keyword}\": {count}");
    }
}
