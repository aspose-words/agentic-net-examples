using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOCX document from disk.
        Document doc = new Document("Input.docx");

        // Iterate through all paragraphs in the first section.
        // For each paragraph, the Bidi property indicates right‑to‑left direction.
        foreach (Paragraph paragraph in doc.FirstSection.Body.Paragraphs)
        {
            bool isRightToLeft = paragraph.ParagraphFormat.Bidi;
            string direction = isRightToLeft ? "Right-to-Left" : "Left-to-Right";

            // Output the detected direction for the current paragraph.
            Console.WriteLine($"Paragraph #{paragraph.GetHashCode()}: {direction}");
        }
    }
}
