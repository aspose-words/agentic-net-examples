using System;
using System.IO;
using System.Text;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOTM file.
        Document doc = new Document("Template.dotm");

        // Access the main story (body) of the first section.
        Body body = doc.FirstSection.Body;

        // Define the range of paragraphs you want to extract (inclusive start, exclusive end).
        // Adjust these indices as needed.
        int startParagraphIndex = 2; // third paragraph (0‑based index)
        int endParagraphIndex   = 5; // up to but not including the sixth paragraph

        // Ensure the indices are within the collection bounds.
        if (startParagraphIndex < 0) startParagraphIndex = 0;
        if (endParagraphIndex > body.Paragraphs.Count) endParagraphIndex = body.Paragraphs.Count;

        // Collect the text of the selected paragraphs.
        StringBuilder extracted = new StringBuilder();

        for (int i = startParagraphIndex; i < endParagraphIndex; i++)
        {
            Paragraph para = body.Paragraphs[i];
            // GetText returns the paragraph text including the paragraph break.
            // Trim the trailing paragraph break characters for cleaner output.
            string text = para.GetText().TrimEnd('\r', '\a');
            extracted.AppendLine(text);
        }

        // Save the extracted content to a plain‑text file.
        File.WriteAllText("ExtractedContent.txt", extracted.ToString());

        Console.WriteLine("Extraction complete. Content saved to ExtractedContent.txt");
    }
}
