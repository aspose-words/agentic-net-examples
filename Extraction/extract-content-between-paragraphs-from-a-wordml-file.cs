using System;
using System.IO;
using System.Text;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the WORDML (WordprocessingML) document from file.
        // The Document constructor handles loading; no custom loading code is needed.
        Document doc = new Document("input.xml");

        // Prepare a StringBuilder to accumulate the text found between paragraphs.
        StringBuilder extractedText = new StringBuilder();

        // Iterate over all paragraphs in the main body of the first section.
        // The Paragraphs collection is provided by the Story.Paragraphs property.
        foreach (Paragraph paragraph in doc.FirstSection.Body.Paragraphs)
        {
            // Get the full text of the paragraph, which includes the paragraph break character.
            string paragraphText = paragraph.GetText();

            // Remove trailing paragraph break characters (\r) to keep only the visible content.
            paragraphText = paragraphText.TrimEnd('\r', '\n');

            // Append the cleaned paragraph text followed by a line break.
            extractedText.AppendLine(paragraphText);
        }

        // Write the extracted content to a plain‑text file.
        // This uses standard .NET I/O; no Aspose.Words saving is required for plain text.
        File.WriteAllText("extracted.txt", extractedText.ToString());
    }
}
