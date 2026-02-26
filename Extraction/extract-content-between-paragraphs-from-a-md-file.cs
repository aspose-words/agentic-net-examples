using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the Markdown file.
        string markdownPath = "input.md";

        // Load the Markdown document while preserving empty lines.
        using (MemoryStream stream = new MemoryStream(File.ReadAllBytes(markdownPath)))
        {
            MarkdownLoadOptions loadOptions = new MarkdownLoadOptions { PreserveEmptyLines = true };
            Document doc = new Document(stream, loadOptions);

            // Access the collection of paragraphs in the main body.
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            // Write each paragraph's content (without the trailing paragraph break) to an output file.
            using (StreamWriter writer = new StreamWriter("extracted.txt"))
            {
                for (int i = 0; i < paragraphs.Count; i++)
                {
                    // Get the full text of the paragraph, then trim the paragraph break characters.
                    string paragraphText = paragraphs[i].GetText().TrimEnd('\r', '\n');

                    writer.WriteLine($"Paragraph {i + 1}:");
                    writer.WriteLine(paragraphText);
                    writer.WriteLine(); // Separate entries with an empty line.
                }
            }
        }
    }
}
