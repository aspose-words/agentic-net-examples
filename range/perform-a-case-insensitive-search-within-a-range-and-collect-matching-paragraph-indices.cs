using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs.
        builder.Writeln("The quick brown fox jumps over the lazy dog.");
        builder.Writeln("A fox is a clever animal.");
        builder.Writeln("This paragraph does not contain the keyword.");
        builder.Writeln("Another line about FOXes in the forest.");
        builder.Writeln("No match here.");

        // Save the document (optional, just to demonstrate saving).
        doc.Save("Sample.docx");

        // The term to search for (case‑insensitive).
        string searchTerm = "fox";

        // Collect indices of paragraphs that contain the search term.
        List<int> matchingParagraphIndices = new List<int>();
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        for (int i = 0; i < paragraphs.Count; i++)
        {
            // Get the text of the paragraph (including the paragraph break).
            string paraText = paragraphs[i].GetText();

            // Perform a case‑insensitive search.
            if (paraText.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0)
                matchingParagraphIndices.Add(i);
        }

        // Output the results.
        Console.WriteLine($"Paragraph indices containing \"{searchTerm}\":");
        foreach (int index in matchingParagraphIndices)
            Console.WriteLine(index);
    }
}
