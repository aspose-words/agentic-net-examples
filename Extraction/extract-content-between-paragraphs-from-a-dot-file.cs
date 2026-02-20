using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the DOT (Word template) file.
        Document doc = new Document("Template.dot");

        // Access the collection of paragraphs in the main body of the first section.
        ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

        // List to hold the text found between each paragraph.
        List<string> extractedTexts = new List<string>();

        // Iterate through all paragraphs.
        for (int i = 0; i < paragraphs.Count; i++)
        {
            // Get the full text of the paragraph (includes the paragraph break).
            string paragraphText = paragraphs[i].GetText();

            // Remove trailing paragraph break characters for cleaner output.
            extractedTexts.Add(paragraphText.TrimEnd('\r', '\n'));
        }

        // Output the extracted texts to the console.
        foreach (string text in extractedTexts)
        {
            Console.WriteLine(text);
        }

        // Save the extracted content to a plain‑text file (optional).
        System.IO.File.WriteAllLines("Extracted.txt", extractedTexts);
    }
}
