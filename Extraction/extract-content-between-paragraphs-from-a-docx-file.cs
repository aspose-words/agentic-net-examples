using System;
using System.Text;
using Aspose.Words;

class ExtractBetweenParagraphs
{
    static void Main()
    {
        // Load the DOCX file (uses the provided load rule)
        Document doc = new Document("Input.docx");

        // Define the indexes of the start and end paragraphs (inclusive).
        // For example, extract content that lies between paragraph 2 and paragraph 5.
        int startParagraphIndex = 2; // zero‑based index of the first boundary paragraph
        int endParagraphIndex   = 5; // zero‑based index of the second boundary paragraph

        // Validate indexes.
        if (startParagraphIndex < 0 || endParagraphIndex >= doc.FirstSection.Body.Paragraphs.Count ||
            startParagraphIndex >= endParagraphIndex)
        {
            Console.WriteLine("Invalid paragraph indexes.");
            return;
        }

        // Build the text that exists strictly between the two boundary paragraphs.
        StringBuilder betweenText = new StringBuilder();

        // Loop over the paragraphs that are between the two boundaries.
        for (int i = startParagraphIndex + 1; i < endParagraphIndex; i++)
        {
            Paragraph para = doc.FirstSection.Body.Paragraphs[i];
            betweenText.Append(para.GetText()); // GetText includes the paragraph break.
        }

        // Output the extracted content.
        Console.WriteLine("Content between paragraph {0} and {1}:", startParagraphIndex, endParagraphIndex);
        Console.WriteLine(betweenText.ToString());
    }
}
