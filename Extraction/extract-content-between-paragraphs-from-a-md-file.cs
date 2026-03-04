using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the Markdown file.
        const string markdownPath = "input.md";

        // Load the Markdown document while preserving empty lines.
        using (FileStream stream = new FileStream(markdownPath, FileMode.Open, FileAccess.Read))
        {
            var loadOptions = new MarkdownLoadOptions { PreserveEmptyLines = true };
            Document doc = new Document(stream, loadOptions);

            // Get all paragraphs in the main body of the document.
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            // Extract the text of each paragraph (excluding the trailing paragraph break).
            var resultBuilder = new StringBuilder();

            for (int i = 0; i < paragraphs.Count; i++)
            {
                // GetText() returns the paragraph text plus a paragraph break character.
                string paragraphText = paragraphs[i].GetText();

                // Remove the trailing paragraph break characters (\r, \n, \f, etc.).
                paragraphText = paragraphText.TrimEnd('\r', '\n', '\f');

                resultBuilder.AppendLine($"Paragraph {i + 1}: {paragraphText}");
            }

            // Output the extracted content.
            Console.WriteLine(resultBuilder.ToString());
        }
    }
}
