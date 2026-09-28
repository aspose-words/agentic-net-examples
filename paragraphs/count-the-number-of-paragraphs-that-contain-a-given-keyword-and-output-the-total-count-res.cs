using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define the keyword to search for.
        string keyword = "keyword";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs to the document.
        builder.Writeln("This is a test paragraph.");
        builder.Writeln("Keyword appears here.");
        builder.Writeln("Another line with the KEYWORD inside.");
        builder.Writeln("No matching word in this one.");

        // Save the sample document (optional, demonstrates lifecycle compliance).
        doc.Save("SampleDocument.docx");

        // Count paragraphs that contain the keyword (case‑insensitive).
        int count = 0;
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Paragraph.GetText() returns the paragraph text including the end‑of‑paragraph marker.
            string text = paragraph.GetText();
            if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                count++;
            }
        }

        // Output the total count to the console.
        Console.WriteLine($"Number of paragraphs containing \"{keyword}\": {count}");
    }
}
