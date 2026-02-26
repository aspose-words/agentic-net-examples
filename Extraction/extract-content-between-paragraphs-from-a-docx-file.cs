using System;
using System.IO;
using System.Text;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX file from disk.
        Document doc = new Document("Input.docx");

        // Example: extract text that lies between paragraph index 2 (inclusive) and 5 (exclusive).
        // Paragraph indices are zero‑based.
        int startIndex = 2;
        int endIndex   = 5;

        // Validate the range.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
        if (startIndex < 0 || endIndex > paragraphs.Count || startIndex >= endIndex)
        {
            Console.WriteLine("Invalid paragraph range.");
            return;
        }

        // Build the extracted text.
        StringBuilder extractedBuilder = new StringBuilder();
        for (int i = startIndex; i < endIndex; i++)
        {
            // GetText returns the paragraph text plus a paragraph break.
            // Trim the trailing break to avoid duplicate empty lines.
            extractedBuilder.Append(paragraphs[i].GetText().TrimEnd('\r', '\n'));
            extractedBuilder.Append(Environment.NewLine);
        }

        string extractedText = extractedBuilder.ToString().TrimEnd();

        // Display the result.
        Console.WriteLine("Extracted content between paragraphs:");
        Console.WriteLine(extractedText);

        // Optionally write the extracted content to a plain‑text file.
        File.WriteAllText("Extracted.txt", extractedText);
    }
}
