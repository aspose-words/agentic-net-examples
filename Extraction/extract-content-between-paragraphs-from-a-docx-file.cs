using System;
using System.IO;
using System.Text;
using Aspose.Words;

class ExtractParagraphContent
{
    static void Main()
    {
        // Load the DOCX file (lifecycle: load)
        Document doc = new Document("Input.docx");

        // StringBuilder to accumulate the text that appears between paragraphs.
        // Each paragraph's text includes its own paragraph break, so we trim it to avoid duplicate breaks.
        StringBuilder betweenText = new StringBuilder();

        // Iterate through all paragraphs in the main story of the document.
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            // Get the raw text of the paragraph (includes the ending paragraph break).
            string paragraphText = para.GetText();

            // Remove the trailing paragraph break to get only the content.
            string trimmed = paragraphText.TrimEnd('\r', '\n');

            // Append the trimmed content followed by a custom separator (e.g., a line).
            betweenText.AppendLine(trimmed);
            betweenText.AppendLine("---"); // separator indicating "between paragraphs"
        }

        // Save the extracted content to a plain‑text file (lifecycle: save)
        File.WriteAllText("ExtractedContent.txt", betweenText.ToString());
    }
}
