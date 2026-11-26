using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs – some contain the keyword, others do not.
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("Keyword appears here.");
        builder.Writeln("Another line without the word.");
        builder.Writeln("Second occurrence of Keyword in this paragraph.");
        builder.Writeln("No keyword here.");

        // Keyword to search for (case‑insensitive).
        string keyword = "Keyword";

        // Count paragraphs that contain the keyword.
        int matchingParagraphs = 0;
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            string text = para.GetText(); // Includes the paragraph mark.
            if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                matchingParagraphs++;
        }

        // Output the total count.
        Console.WriteLine($"Paragraphs containing \"{keyword}\": {matchingParagraphs}");

        // Save the document (optional output file).
        doc.Save("Output.docx");
    }
}
