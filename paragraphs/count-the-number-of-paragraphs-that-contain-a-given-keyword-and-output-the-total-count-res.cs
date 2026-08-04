using System;
using Aspose.Words;

public class ParagraphKeywordCounter
{
    public static void Main()
    {
        // Define the keyword to search for.
        const string keyword = "keyword";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs to the document.
        builder.Writeln("This is the first paragraph without the special word.");
        builder.Writeln("The second paragraph contains the Keyword we are looking for.");
        builder.Writeln("Another line that does not match.");
        builder.Writeln("Keyword appears again in this fourth paragraph.");
        builder.Writeln("Final paragraph without it.");

        // Count paragraphs that contain the keyword (case‑insensitive).
        int matchingParagraphs = 0;
        NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph paragraph in allParagraphs)
        {
            string text = paragraph.GetText(); // Includes the paragraph break.
            if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                matchingParagraphs++;
        }

        // Output the result.
        Console.WriteLine($"Paragraphs containing \"{keyword}\": {matchingParagraphs}");

        // Optional: save the document to verify its contents.
        doc.Save("SampleDocument.docx");
    }
}
