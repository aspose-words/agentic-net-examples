using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Sample paragraphs – some contain the keyword "Aspose", others do not.
        builder.Writeln("This is the first paragraph.");
        builder.Writeln("Aspose.Words is a powerful library for document processing.");
        builder.Writeln("Another paragraph without the keyword.");
        builder.Writeln("Learning Aspose.Words can improve productivity.");
        builder.Writeln("Final paragraph.");

        // Define the keyword to search for.
        string keyword = "Aspose";

        // Count paragraphs that contain the keyword (case‑insensitive).
        int count = 0;
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // Get the paragraph text without the trailing paragraph break.
            string text = para.GetText().TrimEnd('\r', '\a');
            if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                count++;
        }

        // Output the result.
        Console.WriteLine($"Number of paragraphs containing \"{keyword}\": {count}");

        // Save the document (optional, demonstrates document creation).
        doc.Save("ParagraphKeywordCount.docx");
    }
}
