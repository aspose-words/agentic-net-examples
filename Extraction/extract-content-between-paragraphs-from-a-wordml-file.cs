using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the WORDML (WordprocessingML) document from a file.
        // The Document constructor automatically detects the format.
        Document doc = new Document("input.xml");

        // Collect the text of each paragraph in the main body.
        List<string> paragraphTexts = new List<string>();
        foreach (Paragraph para in doc.FirstSection.Body.Paragraphs)
        {
            // GetText() returns the paragraph text including the ending paragraph break.
            // Trim the trailing break characters to obtain clean content.
            string text = para.GetText().TrimEnd('\r', '\n');
            paragraphTexts.Add(text);
        }

        // Example: display the content that lies between each consecutive pair of paragraphs.
        // Here we simply concatenate the two paragraphs with a separator for demonstration.
        for (int i = 0; i < paragraphTexts.Count - 1; i++)
        {
            string between = paragraphTexts[i] + Environment.NewLine + paragraphTexts[i + 1];
            Console.WriteLine($"--- Between paragraph {i} and {i + 1} ---");
            Console.WriteLine(between);
            Console.WriteLine();
        }

        // Optionally, write all extracted paragraph texts to a plain‑text file.
        File.WriteAllLines("extracted_paragraphs.txt", paragraphTexts);
    }
}
