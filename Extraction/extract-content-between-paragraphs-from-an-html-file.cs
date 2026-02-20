using System;
using System.IO;
using Aspose.Words;

class ExtractParagraphContent
{
    static void Main(string[] args)
    {
        // Path to the HTML file. Adjust as needed or pass as a command‑line argument.
        string htmlPath = args.Length > 0 ? args[0] : "input.html";

        // Load the HTML document into an Aspose.Words Document.
        Document doc = new Document(htmlPath);

        // Access the collection of paragraphs in the main story (body) of the first section.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // Iterate through each paragraph and output its plain text.
        for (int i = 0; i < paragraphs.Count; i++)
        {
            // Get the text of the current paragraph (includes the paragraph break).
            string text = paragraphs[i].GetText();

            // Trim the trailing paragraph break characters for cleaner output.
            Console.WriteLine($"Paragraph {i + 1}: {text.TrimEnd('\r', '\n')}");
        }
    }
}
