using System;
using System.IO;
using Aspose.Words;

class ExtractBetweenParagraphs
{
    static void Main()
    {
        // Load the DOTM (macro-enabled template) file.
        // The Document constructor automatically detects the format.
        Document doc = new Document("Template.dotm");

        // Retrieve all paragraph nodes in the document (including those in headers/footers if needed).
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Define the start and end paragraph indices (zero‑based).
        // Adjust these values to the desired range.
        int startIndex = 2; // third paragraph
        int endIndex = 5;   // sixth paragraph

        // Validate indices.
        if (startIndex < 0 || endIndex >= paragraphs.Count || startIndex > endIndex)
        {
            Console.WriteLine("Invalid paragraph range specified.");
            return;
        }

        // Accumulate the text of the selected paragraphs.
        StringWriter extractedText = new StringWriter();
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = (Paragraph)paragraphs[i];
            // GetText() returns the paragraph text including the end‑of‑paragraph marker.
            extractedText.Write(para.GetText());
        }

        // Output the extracted content to the console.
        Console.WriteLine("Extracted text between paragraphs {0} and {1}:", startIndex + 1, endIndex + 1);
        Console.WriteLine(extractedText.ToString());

        // Optionally, save the extracted content to a plain‑text file.
        File.WriteAllText("ExtractedContent.txt", extractedText.ToString());
    }
}
