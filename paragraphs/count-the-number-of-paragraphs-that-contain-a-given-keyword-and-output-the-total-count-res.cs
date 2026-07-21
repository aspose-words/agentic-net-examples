using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // The keyword we want to search for in paragraphs.
        const string keyword = "Aspose";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs – some contain the keyword, others do not.
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("Aspose.Words is a powerful library for document processing.");
        builder.Writeln("Another paragraph without the target word.");
        builder.Writeln("Learning Aspose can improve productivity.");
        builder.Writeln("Final paragraph.");

        // Count paragraphs that contain the keyword (case‑insensitive).
        int count = 0;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // GetText() returns the paragraph text including the end‑of‑paragraph mark.
            // Trim it to avoid false negatives due to trailing whitespace.
            string text = para.GetText().Trim();

            if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                count++;
        }

        // Output the result.
        Console.WriteLine($"Number of paragraphs containing \"{keyword}\": {count}");

        // Save the document so the example is self‑contained.
        doc.Save("ParagraphKeywordCount.docx");
    }
}
