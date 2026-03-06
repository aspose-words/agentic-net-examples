using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the Markdown file.
        const string markdownPath = "input.md";

        // Load the Markdown document while preserving empty lines.
        var loadOptions = new MarkdownLoadOptions { PreserveEmptyLines = true };
        Document doc;
        using (FileStream stream = new FileStream(markdownPath, FileMode.Open, FileAccess.Read))
        {
            doc = new Document(stream, loadOptions);
        }

        // Access the collection of paragraphs in the main body.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Iterate through consecutive paragraph pairs and output the text that lies between them.
        // In Aspose.Words the "content between paragraphs" is represented by the paragraphs themselves,
        // so we display the text of each paragraph pair.
        for (int i = 0; i < paragraphs.Count - 1; i++)
        {
            // Get raw text of each paragraph (includes the paragraph break character).
            string firstParagraph = paragraphs[i].GetText().TrimEnd('\r', '\n');
            string secondParagraph = paragraphs[i + 1].GetText().TrimEnd('\r', '\n');

            Console.WriteLine($"--- Between Paragraph {i} and {i + 1} ---");
            Console.WriteLine($"First : {firstParagraph}");
            Console.WriteLine($"Second: {secondParagraph}");
            Console.WriteLine();
        }
    }
}
